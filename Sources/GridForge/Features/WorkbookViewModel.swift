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

    init() {
        let wb = Workbook()
        self.workbook = wb
        self.selection = SelectionState()
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
        undoManager.perform(command, on: workbook)

        // Recalculate formulas
        formulaEngine.recalculateAffected(changedCell: address, in: sheet)

        syncEditingTextFromActiveCell()
        markDirty()
        bump()
    }

    func cancelEdit() {
        isEditing = false
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
        undoManager.perform(command, on: workbook)

        // Recalculate formulas for all cleared addresses
        for addr in selectedRange.allAddresses {
            formulaEngine.recalculateAffected(changedCell: addr, in: activeSheet)
        }

        syncEditingTextFromActiveCell()
        markDirty()
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
        undoManager.perform(command, on: workbook)
        formulaEngine.recalculateAffected(changedCell: address, in: sheet)
        markDirty()
        bump()
    }

    // MARK: - Batch Operations

    func batchSetCellRawInputs(_ changes: [(CellAddress, String)]) {
        guard !changes.isEmpty else { return }

        let sheet = activeSheet
        var commands: [SpreadsheetCommand] = []

        for (address, rawInput) in changes {
            let oldCell = sheet.cell(at: address)
            let oldRawInput = oldCell?.rawInput ?? ""
            let oldValue = oldCell?.value ?? .empty

            let cmd = SetCellValueCommand(
                sheetIndex: workbook.activeSheetIndex,
                address: address,
                newRawInput: rawInput,
                oldRawInput: oldRawInput,
                oldValue: oldValue
            )
            commands.append(cmd)
        }

        let compound = CompoundCommand(commands: commands, label: "Batch Edit (\(changes.count) cells)")
        undoManager.perform(compound, on: workbook)

        // Single recalculation pass
        formulaEngine.recalculate(worksheet: sheet)

        markDirty()
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

        batchSetCellRawInputs(changes)
    }

    // MARK: - Undo/Redo

    func undo() {
        if isEditing { cancelEdit() }
        if undoManager.undo(on: workbook) {
            formulaEngine.recalculate(worksheet: activeSheet)
            syncEditingTextFromActiveCell()
            markDirty()
            bump()
        }
    }

    func redo() {
        if isEditing { cancelEdit() }
        if undoManager.redo(on: workbook) {
            formulaEngine.recalculate(worksheet: activeSheet)
            syncEditingTextFromActiveCell()
            markDirty()
            bump()
        }
    }

    // MARK: - Sheet management

    func addSheet() {
        workbook.addSheet()
        markDirty()
        bump()
    }

    func deleteSheet(at index: Int) {
        guard workbook.sheets.count > 1 else { return }
        workbook.deleteSheet(at: index)
        selection = SelectionState()
        syncEditingTextFromActiveCell()
        formulaEngine.recalculate(worksheet: activeSheet)
        markDirty()
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
        workbook.renameSheet(at: index, to: name)
        markDirty()
        bump()
    }

    func duplicateSheet(at index: Int) {
        workbook.duplicateSheet(at: index)
        markDirty()
        bump()
    }

    // MARK: - Row/Column operations

    func insertRow() {
        if isEditing { commitEdit() }
        let command = InsertRowCommand(
            sheetIndex: workbook.activeSheetIndex,
            rowIndex: activeCell.row
        )
        undoManager.perform(command, on: workbook)
        formulaEngine.recalculate(worksheet: activeSheet)
        markDirty()
        bump()
    }

    func insertColumn() {
        if isEditing { commitEdit() }
        let command = InsertColumnCommand(
            sheetIndex: workbook.activeSheetIndex,
            columnIndex: activeCell.column
        )
        undoManager.perform(command, on: workbook)
        formulaEngine.recalculate(worksheet: activeSheet)
        markDirty()
        bump()
    }

    func deleteRow() {
        if isEditing { commitEdit() }
        let command = DeleteRowCommand(
            sheetIndex: workbook.activeSheetIndex,
            rowIndex: activeCell.row
        )
        undoManager.perform(command, on: workbook)
        formulaEngine.recalculate(worksheet: activeSheet)
        markDirty()
        bump()
    }

    func deleteColumn() {
        if isEditing { commitEdit() }
        let command = DeleteColumnCommand(
            sheetIndex: workbook.activeSheetIndex,
            columnIndex: activeCell.column
        )
        undoManager.perform(command, on: workbook)
        formulaEngine.recalculate(worksheet: activeSheet)
        markDirty()
        bump()
    }

    // MARK: - Formatting

    func toggleBold() {
        let cell = activeSheet.cell(at: activeCell) ?? {
            let c = Cell()
            activeSheet.cells[activeCell] = c
            return c
        }()
        cell.formatting.bold.toggle()
        markDirty()
        bump()
    }

    func toggleItalic() {
        let cell = activeSheet.cell(at: activeCell) ?? {
            let c = Cell()
            activeSheet.cells[activeCell] = c
            return c
        }()
        cell.formatting.italic.toggle()
        markDirty()
        bump()
    }

    func toggleUnderline() {
        let cell = activeSheet.cell(at: activeCell) ?? {
            let c = Cell()
            activeSheet.cells[activeCell] = c
            return c
        }()
        cell.formatting.underline.toggle()
        markDirty()
        bump()
    }

    func setFontSize(_ size: Double) {
        let clamped = min(max(size, 11), 72)
        let cell = activeSheet.cell(at: activeCell) ?? {
            let c = Cell()
            activeSheet.cells[activeCell] = c
            return c
        }()
        cell.formatting.fontSize = clamped
        markDirty()
        bump()
    }

    func setAlignment(_ alignment: GridForgeCore.HorizontalAlignment) {
        let cell = activeSheet.cell(at: activeCell) ?? {
            let c = Cell()
            activeSheet.cells[activeCell] = c
            return c
        }()
        cell.formatting.alignment = alignment
        markDirty()
        bump()
    }

    // MARK: - Column width

    func setColumnWidth(_ column: Int, _ width: Double) {
        activeSheet.setColumnWidth(width, for: column)
        markDirty()
        bump()
    }

    // MARK: - File operations

    func newWorkbook() {
        workbook = Workbook()
        selection = SelectionState()
        isEditing = false
        editingText = ""
        undoManager.clear()
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

    private func bump() {
        version += 1
    }

    private func markDirty() {
        isDirty = true
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
