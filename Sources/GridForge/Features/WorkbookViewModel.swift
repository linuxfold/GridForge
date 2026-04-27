import SwiftUI
import AppKit
import GridForgeCore
import UniformTypeIdentifiers

// MARK: - FindScope

enum FindScope {
    case activeSheet
    case allSheets
}

// MARK: - WorkbookViewModel

final class WorkbookViewModel: ObservableObject {

    // MARK: Published State

    @Published var workbook: Workbook
    @Published var selection: SelectionState
    @Published var isEditing: Bool = false
    @Published var editingText: String = ""
    @Published var formulaBarHasFocus: Bool = false
    @Published var showInspector: Bool = false
    @Published var showFormulaBar: Bool = true
    @Published var showSheetTabs: Bool = true
    @Published var showStatusBar: Bool = true
    @Published var version: Int = 0

    // Error / Loading state
    @Published var lastError: String?
    @Published var isLoading: Bool = false

    // Dirty flag & current file
    @Published var isDirty: Bool = false
    @Published var currentFileURL: URL?

    // Zoom
    @Published var zoomLevel: Double = 1.0

    // Scroll target (set by scrollToCell, consumed by grid view)
    @Published var scrollTarget: CellAddress?

    // MARK: Internal engines

    let formulaEngine = FormulaEngine()
    let undoManager = SpreadsheetUndoManager()

    // MARK: Display grid size (virtual)

    let displayColumns: Int = 52   // A..AZ
    let displayRows: Int = 1000

    // MARK: Computed convenience

    var activeSheet: Worksheet { workbook.activeSheet }
    var activeCell: CellAddress { selection.activeCell }
    var selectedRange: CellRange { selection.selectedRange }

    // MARK: Computed - Window title

    var windowTitle: String {
        let name: String
        if let url = currentFileURL {
            name = url.deletingPathExtension().lastPathComponent
        } else {
            name = "Untitled"
        }
        return isDirty ? "\(name) \u{2014} Edited" : name
    }

    // MARK: Computed - Undo/Redo state

    var canUndo: Bool { undoManager.canUndo }
    var canRedo: Bool { undoManager.canRedo }

    // MARK: Computed - Status bar

    var statusText: String {
        if isLoading { return "Loading..." }
        if let error = lastError { return error }
        if isEditing { return "Editing \(activeCell.displayString)" }
        return "Ready"
    }

    var cellCountText: String {
        let range = selectedRange
        if range.isSingleCell { return "" }
        return "Cells: \(range.cellCount)"
    }

    var selectionSummary: String {
        let range = selectedRange
        guard !range.isSingleCell else { return "" }
        let sheet = activeSheet

        var numbers: [Double] = []
        var nonEmptyCount = 0
        for addr in range.allAddresses {
            guard let cell = sheet.cell(at: addr) else { continue }
            if cell.isEmpty { continue }
            nonEmptyCount += 1
            if let n = cell.value.numericValue, !cell.value.isEmpty {
                numbers.append(n)
            }
        }

        guard !numbers.isEmpty else {
            if nonEmptyCount > 0 {
                return "Count: \(nonEmptyCount)"
            }
            return ""
        }

        let sum = numbers.reduce(0, +)
        let avg = sum / Double(numbers.count)

        let formatter = NumberFormatter()
        formatter.numberStyle = .decimal
        formatter.maximumFractionDigits = 2

        let sumStr = formatter.string(from: NSNumber(value: sum)) ?? String(sum)
        let avgStr = formatter.string(from: NSNumber(value: avg)) ?? String(avg)

        return "Sum: \(sumStr) | Average: \(avgStr) | Count: \(numbers.count)"
    }

    // MARK: Init

    init(workbook: Workbook = Workbook()) {
        self.workbook = workbook
        self.selection = SelectionState()
        self.undoManager.markClean()
        syncEditingTextFromActiveCell()
    }

    // MARK: - Selection

    func selectCell(_ address: CellAddress) {
        if isEditing {
            commitEdit()
        }
        selection.select(cell: address)
        syncEditingTextFromActiveCell()
        bump()
    }

    // MARK: - Editing

    func startEditing(withText text: String? = nil) {
        isEditing = true
        if let text = text {
            editingText = text
        } else {
            let cell = activeSheet.cell(at: activeCell)
            editingText = cell?.editString ?? ""
        }
    }

    func commitEdit() {
        guard isEditing else { return }
        isEditing = false
        formulaBarHasFocus = false
        editingText = autoCompleteFormulaIfPossible(editingText)

        let address = activeCell
        let sheet = activeSheet
        let oldCell = sheet.cell(at: address)
        let oldRawInput = oldCell?.rawInput ?? ""
        let oldValue = oldCell?.value ?? .empty

        let command = SetCellValueCommand(
            sheetIndex: workbook.activeSheetIndex,
            address: address,
            newRawInput: editingText,
            oldRawInput: oldRawInput,
            oldValue: oldValue
        )
        perform(command)

        // Recalculate formulas
        formulaEngine.recalculateAffected(changedCell: address, in: sheet)

        syncEditingTextFromActiveCell()
        bump()
    }

    func cancelEdit() {
        isEditing = false
        formulaBarHasFocus = false
        syncEditingTextFromActiveCell()
    }

    // MARK: - Navigation

    func moveSelection(direction: Direction, extend: Bool) {
        if isEditing && !extend {
            commitEdit()
        }
        selection.moveActiveCell(
            direction: direction,
            extend: extend,
            maxColumn: displayColumns,
            maxRow: displayRows
        )
        syncEditingTextFromActiveCell()
        bump()
    }

    // MARK: - Cell operations

    func deleteSelectedCells() {
        if isEditing { cancelEdit() }

        let command = ClearRangeCommand(
            sheetIndex: workbook.activeSheetIndex,
            range: selectedRange
        )
        perform(command)

        // Recalculate formulas for all cleared addresses
        for addr in selectedRange.allAddresses {
            formulaEngine.recalculateAffected(changedCell: addr, in: activeSheet)
        }

        syncEditingTextFromActiveCell()
        bump()
    }

    func setCellRawInput(at address: CellAddress, rawInput: String) {
        let sheet = activeSheet
        let oldCell = sheet.cell(at: address)
        let oldRawInput = oldCell?.rawInput ?? ""
        let oldValue = oldCell?.value ?? .empty

        let command = SetCellValueCommand(
            sheetIndex: workbook.activeSheetIndex,
            address: address,
            newRawInput: rawInput,
            oldRawInput: oldRawInput,
            oldValue: oldValue
        )
        perform(command)
        formulaEngine.recalculateAffected(changedCell: address, in: sheet)
        bump()
    }

    // MARK: - Batch Operations

    func batchSetCellRawInputs(_ changes: [(CellAddress, String)], label: String? = nil) {
        guard !changes.isEmpty else { return }

        let sheet = activeSheet
        let command = SetCellValuesCommand(
            sheetIndex: workbook.activeSheetIndex,
            changes: changes,
            label: label ?? "Batch Edit (\(changes.count) cells)"
        )
        perform(command)

        // Single recalculation pass
        formulaEngine.recalculate(worksheet: sheet)

        bump()
    }

    // MARK: - Clipboard

    func copy() {
        let range = selectedRange
        let sheet = activeSheet

        var rows: [String] = []
        for r in range.start.row ... range.end.row {
            var cols: [String] = []
            for c in range.start.column ... range.end.column {
                let addr = CellAddress(column: c, row: r)
                let cell = sheet.cell(at: addr)
                // Copy rawInput for formula cells so formulas are preserved
                let text: String
                if let cell = cell, cell.isFormula {
                    text = cell.rawInput
                } else {
                    text = cell?.displayString ?? ""
                }
                cols.append(text)
            }
            rows.append(cols.joined(separator: "\t"))
        }
        let text = rows.joined(separator: "\n")

        let pb = NSPasteboard.general
        pb.clearContents()
        pb.setString(text, forType: .string)
    }

    func cut() {
        copy()
        deleteSelectedCells()
    }

    func paste() {
        guard let text = NSPasteboard.general.string(forType: .string) else { return }
        if isEditing { commitEdit() }

        let rows = text.components(separatedBy: .newlines)
        let startRow = activeCell.row
        let startCol = activeCell.column

        var changes: [(CellAddress, String)] = []
        for (rowOffset, rowText) in rows.enumerated() {
            let cols = rowText.components(separatedBy: "\t")
            for (colOffset, cellText) in cols.enumerated() {
                let addr = CellAddress(column: startCol + colOffset, row: startRow + rowOffset)
                changes.append((addr, cellText))
            }
        }

        batchSetCellRawInputs(changes, label: "Paste (\(changes.count) cells)")
    }

    // MARK: - Undo/Redo

    func undo() {
        if isEditing { cancelEdit() }
        if undoManager.undo(on: workbook) {
            formulaEngine.recalculate(worksheet: activeSheet)
            syncEditingTextFromActiveCell()
            syncDirtyState()
            bump()
        }
    }

    func redo() {
        if isEditing { cancelEdit() }
        if undoManager.redo(on: workbook) {
            formulaEngine.recalculate(worksheet: activeSheet)
            syncEditingTextFromActiveCell()
            syncDirtyState()
            bump()
        }
    }

    // MARK: - Sheet management

    func addSheet() {
        perform(AddSheetCommand())
        bump()
    }

    func deleteSheet(at index: Int) {
        guard workbook.sheets.count > 1 else { return }
        perform(DeleteSheetCommand(sheetIndex: index))
        selection = SelectionState()
        syncEditingTextFromActiveCell()
        formulaEngine.recalculate(worksheet: activeSheet)
        bump()
    }

    func switchSheet(to index: Int) {
        guard index >= 0, index < workbook.sheets.count else { return }
        if isEditing { commitEdit() }
        workbook.activeSheetIndex = index
        selection = SelectionState()
        syncEditingTextFromActiveCell()
        bump()
    }

    func renameSheet(at index: Int, to name: String) {
        guard index >= 0, index < workbook.sheets.count else { return }
        let trimmed = name.trimmingCharacters(in: .whitespaces)
        guard !trimmed.isEmpty, trimmed != workbook.sheets[index].name else { return }
        perform(RenameSheetCommand(sheetIndex: index, oldName: workbook.sheets[index].name, newName: trimmed))
        bump()
    }

    func duplicateSheet(at index: Int) {
        guard index >= 0, index < workbook.sheets.count else { return }
        perform(DuplicateSheetCommand(sourceIndex: index))
        bump()
    }

    func moveSheet(from sourceIndex: Int, to destinationIndex: Int) {
        guard sourceIndex >= 0,
              sourceIndex < workbook.sheets.count,
              destinationIndex >= 0,
              destinationIndex < workbook.sheets.count,
              sourceIndex != destinationIndex else { return }
        perform(MoveSheetCommand(fromIndex: sourceIndex, toIndex: destinationIndex))
        bump()
    }

    // MARK: - Row/Column operations

    func insertRow() {
        if isEditing { commitEdit() }
        let command = InsertRowCommand(
            sheetIndex: workbook.activeSheetIndex,
            rowIndex: activeCell.row
        )
        perform(command)
        formulaEngine.recalculate(worksheet: activeSheet)
        bump()
    }

    func insertColumn() {
        if isEditing { commitEdit() }
        let command = InsertColumnCommand(
            sheetIndex: workbook.activeSheetIndex,
            columnIndex: activeCell.column
        )
        perform(command)
        formulaEngine.recalculate(worksheet: activeSheet)
        bump()
    }

    func deleteRow() {
        if isEditing { commitEdit() }
        let command = DeleteRowCommand(
            sheetIndex: workbook.activeSheetIndex,
            rowIndex: activeCell.row
        )
        perform(command)
        formulaEngine.recalculate(worksheet: activeSheet)
        bump()
    }

    func deleteColumn() {
        if isEditing { commitEdit() }
        let command = DeleteColumnCommand(
            sheetIndex: workbook.activeSheetIndex,
            columnIndex: activeCell.column
        )
        perform(command)
        formulaEngine.recalculate(worksheet: activeSheet)
        bump()
    }

    // MARK: - Formatting

    func toggleBold() {
        updateFormatting(label: "Toggle Bold") { $0.bold.toggle() }
    }

    func toggleItalic() {
        updateFormatting(label: "Toggle Italic") { $0.italic.toggle() }
    }

    func toggleUnderline() {
        updateFormatting(label: "Toggle Underline") { $0.underline.toggle() }
    }

    func setFontSize(_ size: Double) {
        let clamped = min(max(size, 11), 72)
        updateFormatting(label: "Set Font Size") { $0.fontSize = clamped }
    }

    func setAlignment(_ alignment: GridForgeCore.HorizontalAlignment) {
        updateFormatting(label: "Set Alignment") { $0.alignment = alignment }
    }

    // MARK: - Column width

    func setColumnWidth(_ column: Int, _ width: Double) {
        let oldWidth = activeSheet.columnWidth(for: column)
        guard abs(oldWidth - width) > 0.25 else { return }
        perform(ResizeColumnCommand(
            sheetIndex: workbook.activeSheetIndex,
            columnIndex: column,
            oldWidth: oldWidth,
            newWidth: width
        ))
        bump()
    }

    // MARK: - File operations

    func newWorkbook() {
        workbook = Workbook()
        selection = SelectionState()
        isEditing = false
        editingText = ""
        undoManager.clear()
        undoManager.markClean()
        formulaEngine.dependencyGraph.clear()
        isDirty = false
        currentFileURL = nil
        bump()
    }

    func openFile(url: URL) {
        isLoading = true
        bump()
        do {
            let wb = try XLSXReader.read(from: url)
            workbook = wb
            selection = SelectionState()
            isEditing = false
            editingText = ""
            undoManager.clear()
            undoManager.markClean()
            formulaEngine.recalculate(worksheet: activeSheet)
            syncEditingTextFromActiveCell()
            isDirty = false
            currentFileURL = url
            isLoading = false
            bump()
        } catch {
            isLoading = false
            setError("Failed to open file: \(error.localizedDescription)")
            bump()
        }
    }

    func saveFile(url: URL) {
        isLoading = true
        bump()
        do {
            try XLSXWriter.write(workbook, to: url)
            undoManager.markClean()
            isDirty = false
            currentFileURL = url
            isLoading = false
            bump()
        } catch {
            isLoading = false
            setError("Failed to save file: \(error.localizedDescription)")
            bump()
        }
    }

    func revertToSaved() {
        guard let url = currentFileURL else { return }
        openFile(url: url)
    }

    // MARK: - Select All

    func selectAll() {
        let sheet = activeSheet
        if let usedRange = sheet.usedRange {
            let maxCol = max(usedRange.end.column, displayColumns - 1)
            let maxRow = max(usedRange.end.row, displayRows - 1)
            selection.activeCell = CellAddress(column: 0, row: 0)
            selection.selectedRange = CellRange(
                start: CellAddress(column: 0, row: 0),
                end: CellAddress(column: maxCol, row: maxRow)
            )
        } else {
            selection.activeCell = CellAddress(column: 0, row: 0)
            selection.selectedRange = CellRange(
                start: CellAddress(column: 0, row: 0),
                end: CellAddress(column: displayColumns - 1, row: displayRows - 1)
            )
        }
        bump()
    }

    // MARK: - Scroll to cell

    func scrollToCell(_ address: CellAddress) {
        selectCell(address)
        scrollTarget = address
    }

    // MARK: - Find and Replace

    func findAndReplace(find searchText: String, replace replaceText: String?, in scope: FindScope) {
        guard !searchText.isEmpty else { return }

        let sheets: [Worksheet]
        switch scope {
        case .activeSheet:
            sheets = [activeSheet]
        case .allSheets:
            sheets = workbook.sheets
        }

        if let replaceText = replaceText {
            // Find and replace
            var changes: [(CellAddress, String)] = []
            for sheet in sheets {
                for (addr, cell) in sheet.cells {
                    if cell.rawInput.localizedCaseInsensitiveContains(searchText) {
                        let newInput = cell.rawInput.replacingOccurrences(
                            of: searchText,
                            with: replaceText,
                            options: .caseInsensitive
                        )
                        changes.append((addr, newInput))
                    }
                }
            }
            if !changes.isEmpty {
                batchSetCellRawInputs(changes)
            }
        } else {
            // Find only: select the first match
            for sheet in sheets {
                for (addr, cell) in sheet.cells {
                    if cell.rawInput.localizedCaseInsensitiveContains(searchText) ||
                       cell.displayString.localizedCaseInsensitiveContains(searchText) {
                        scrollToCell(addr)
                        return
                    }
                }
            }
        }
    }

    // MARK: - CSV Export

    func exportCSV(to url: URL) {
        isLoading = true
        bump()
        do {
            let sheet = activeSheet
            guard let usedRange = sheet.usedRange else {
                try "".write(to: url, atomically: true, encoding: .utf8)
                isLoading = false
                bump()
                return
            }

            var csvRows: [String] = []
            for r in usedRange.start.row ... usedRange.end.row {
                var cols: [String] = []
                for c in usedRange.start.column ... usedRange.end.column {
                    let addr = CellAddress(column: c, row: r)
                    let display = sheet.cell(at: addr)?.displayString ?? ""
                    // Escape CSV: quote if contains comma, newline or quote
                    if display.contains(",") || display.contains("\n") || display.contains("\"") {
                        let escaped = display.replacingOccurrences(of: "\"", with: "\"\"")
                        cols.append("\"\(escaped)\"")
                    } else {
                        cols.append(display)
                    }
                }
                csvRows.append(cols.joined(separator: ","))
            }

            let csvString = csvRows.joined(separator: "\n")
            try csvString.write(to: url, atomically: true, encoding: .utf8)
            isLoading = false
            bump()
        } catch {
            isLoading = false
            setError("Failed to export CSV: \(error.localizedDescription)")
            bump()
        }
    }

    // MARK: - Zoom

    func zoomIn() {
        zoomLevel = min(zoomLevel + 0.1, 3.0)
        bump()
    }

    func zoomOut() {
        zoomLevel = max(zoomLevel - 0.1, 0.25)
        bump()
    }

    func zoomActualSize() {
        zoomLevel = 1.0
        bump()
    }

    // MARK: - Private helpers

    private func syncEditingTextFromActiveCell() {
        let cell = activeSheet.cell(at: activeCell)
        editingText = cell?.editString ?? ""
    }

    private func autoCompleteFormulaIfPossible(_ input: String) -> String {
        guard input.hasPrefix("="),
              unmatchedOpeningParentheses(in: input) == 1 else {
            return input
        }

        let candidate = input + ")"
        guard formulaParses(candidate) else { return input }
        return candidate
    }

    private func unmatchedOpeningParentheses(in text: String) -> Int? {
        var balance = 0
        var inString = false

        for character in text {
            if character == "\"" {
                inString.toggle()
                continue
            }
            guard !inString else { continue }

            if character == "(" {
                balance += 1
            } else if character == ")" {
                balance -= 1
                if balance < 0 { return nil }
            }
        }

        return inString ? nil : balance
    }

    private func formulaParses(_ input: String) -> Bool {
        let formula = String(input.dropFirst()).trimmingCharacters(in: .whitespacesAndNewlines)
        guard !formula.isEmpty else { return false }

        do {
            let tokens = try Tokenizer(formula: formula).tokenize()
            _ = try FormulaParser(tokens: tokens).parse()
            return true
        } catch {
            return false
        }
    }

    private func perform(_ command: SpreadsheetCommand) {
        undoManager.perform(command, on: workbook)
        syncDirtyState()
    }

    private func updateFormatting(label: String, mutate: (inout CellFormatting) -> Void) {
        let range = selectedRange
        var newFormatting: [CellAddress: CellFormatting] = [:]
        for address in range.allAddresses {
            var formatting = activeSheet.cell(at: address)?.formatting ?? CellFormatting()
            mutate(&formatting)
            newFormatting[address] = formatting
        }

        perform(FormatRangeCommand(
            sheetIndex: workbook.activeSheetIndex,
            range: range,
            newFormatting: newFormatting,
            label: label
        ))
        bump()
    }

    private func bump() {
        version += 1
    }

    private func syncDirtyState() {
        isDirty = undoManager.isDirty
    }

    private func markDirty() {
        syncDirtyState()
    }

    private func setError(_ message: String) {
        lastError = message
        // Auto-clear after 5 seconds
        DispatchQueue.main.asyncAfter(deadline: .now() + 5) { [weak self] in
            if self?.lastError == message {
                self?.lastError = nil
            }
        }
    }
}
