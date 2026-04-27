import Foundation

// MARK: - Command Protocol

/// A reversible command that can be executed and undone on a workbook
public protocol SpreadsheetCommand {
    /// Execute the command (apply the change)
    func execute(on workbook: Workbook)
    /// Reverse the command (restore prior state)
    func undo(on workbook: Workbook)
    /// Human-readable description of the command
    var description: String { get }
}

private struct CommandRecord {
    let revision: Int
    let command: SpreadsheetCommand
}

// MARK: - Undo Manager

/// Command-pattern undo/redo manager for spreadsheet operations
public class SpreadsheetUndoManager {
    private var undoStack: [CommandRecord] = []
    private var redoStack: [CommandRecord] = []
    private var nextRevision: Int = 1
    private var currentRevision: Int = 0
    private var cleanRevision: Int = 0

    public init() {}

    /// Execute a command and push it onto the undo stack.
    /// Clears the redo stack (new action invalidates redo history).
    public func perform(_ command: SpreadsheetCommand, on workbook: Workbook) {
        command.execute(on: workbook)
        let revision = nextRevision
        nextRevision += 1
        undoStack.append(CommandRecord(revision: revision, command: command))
        currentRevision = revision
        redoStack.removeAll()
    }

    /// Undo the most recent command. Returns true if an undo was performed.
    @discardableResult
    public func undo(on workbook: Workbook) -> Bool {
        guard let record = undoStack.popLast() else { return false }
        record.command.undo(on: workbook)
        redoStack.append(record)
        currentRevision = undoStack.last?.revision ?? 0
        return true
    }

    /// Redo the most recently undone command. Returns true if a redo was performed.
    @discardableResult
    public func redo(on workbook: Workbook) -> Bool {
        guard let record = redoStack.popLast() else { return false }
        record.command.execute(on: workbook)
        undoStack.append(record)
        currentRevision = record.revision
        return true
    }

    /// Whether there are commands available to undo
    public var canUndo: Bool { !undoStack.isEmpty }

    /// Whether there are commands available to redo
    public var canRedo: Bool { !redoStack.isEmpty }

    /// Whether the command cursor differs from the last saved/opened state.
    public var isDirty: Bool { currentRevision != cleanRevision }

    /// Mark the current command cursor as clean after a successful save/open.
    public func markClean() {
        cleanRevision = currentRevision
    }

    /// Clear all undo/redo history
    public func clear() {
        undoStack.removeAll()
        redoStack.removeAll()
        nextRevision = 1
        currentRevision = 0
        cleanRevision = 0
    }

    /// The description of the next command to undo, if any
    public var undoDescription: String? {
        undoStack.last?.command.description
    }

    /// The description of the next command to redo, if any
    public var redoDescription: String? {
        redoStack.last?.command.description
    }
}

// MARK: - Concrete Commands

/// Sets a single cell's raw input value, storing old values for undo
public class SetCellValueCommand: SpreadsheetCommand {
    public let sheetIndex: Int
    public let address: CellAddress
    public let newRawInput: String
    public let oldRawInput: String
    public let oldValue: CellValue

    public init(
        sheetIndex: Int,
        address: CellAddress,
        newRawInput: String,
        oldRawInput: String,
        oldValue: CellValue
    ) {
        self.sheetIndex = sheetIndex
        self.address = address
        self.newRawInput = newRawInput
        self.oldRawInput = oldRawInput
        self.oldValue = oldValue
    }

    public func execute(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        let sheet = workbook.sheets[sheetIndex]
        sheet.setRawInput(at: address, rawInput: newRawInput)
    }

    public func undo(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        let sheet = workbook.sheets[sheetIndex]
        if oldRawInput.isEmpty {
            sheet.clearCell(at: address)
        } else {
            let cell = sheet.setRawInput(at: address, rawInput: oldRawInput)
            // Restore the exact old value (important for formulas whose result may differ)
            cell.value = oldValue
        }
    }

    public var description: String { "Set \(address.displayString)" }
}

/// Sets many cell raw inputs as one undoable mutation (paste, find/replace, import patches).
public class SetCellValuesCommand: SpreadsheetCommand {
    public let sheetIndex: Int
    public let changes: [(CellAddress, String)]
    public let label: String
    private var savedCells: [CellAddress: Cell] = [:]
    private var missingCells = Set<CellAddress>()

    public init(sheetIndex: Int, changes: [(CellAddress, String)], label: String = "Batch Edit") {
        self.sheetIndex = sheetIndex
        self.changes = changes
        self.label = label
    }

    public func execute(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        let sheet = workbook.sheets[sheetIndex]
        savedCells = [:]
        missingCells = []
        for (address, rawInput) in changes {
            if let existing = sheet.cells[address] {
                savedCells[address] = existing.copy()
            } else {
                missingCells.insert(address)
            }
            sheet.setRawInput(at: address, rawInput: rawInput)
        }
    }

    public func undo(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        let sheet = workbook.sheets[sheetIndex]
        for address in missingCells {
            sheet.cells.removeValue(forKey: address)
        }
        for (address, cell) in savedCells {
            sheet.cells[address] = cell
        }
    }

    public var description: String { label }
}

/// Replaces formatting for a range as one reversible mutation.
public class FormatRangeCommand: SpreadsheetCommand {
    public let sheetIndex: Int
    public let range: CellRange
    public let newFormatting: [CellAddress: CellFormatting]
    public let label: String
    private var savedCells: [CellAddress: Cell] = [:]
    private var missingCells = Set<CellAddress>()

    public init(
        sheetIndex: Int,
        range: CellRange,
        newFormatting: [CellAddress: CellFormatting],
        label: String = "Format Range"
    ) {
        self.sheetIndex = sheetIndex
        self.range = range
        self.newFormatting = newFormatting
        self.label = label
    }

    public func execute(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        let sheet = workbook.sheets[sheetIndex]
        savedCells = [:]
        missingCells = []
        for address in range.allAddresses {
            let cell: Cell
            if let existing = sheet.cells[address] {
                savedCells[address] = existing.copy()
                cell = existing
            } else {
                missingCells.insert(address)
                cell = Cell()
                sheet.cells[address] = cell
            }
            if let formatting = newFormatting[address] {
                cell.formatting = formatting
            }
        }
    }

    public func undo(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        let sheet = workbook.sheets[sheetIndex]
        for address in missingCells {
            sheet.cells.removeValue(forKey: address)
        }
        for (address, cell) in savedCells {
            sheet.cells[address] = cell
        }
    }

    public var description: String { label }
}

/// Resizes a column and restores the previous width on undo.
public class ResizeColumnCommand: SpreadsheetCommand {
    public let sheetIndex: Int
    public let columnIndex: Int
    public let oldWidth: Double
    public let newWidth: Double

    public init(sheetIndex: Int, columnIndex: Int, oldWidth: Double, newWidth: Double) {
        self.sheetIndex = sheetIndex
        self.columnIndex = columnIndex
        self.oldWidth = oldWidth
        self.newWidth = newWidth
    }

    public func execute(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        workbook.sheets[sheetIndex].setColumnWidth(newWidth, for: columnIndex)
    }

    public func undo(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        workbook.sheets[sheetIndex].setColumnWidth(oldWidth, for: columnIndex)
    }

    public var description: String {
        "Resize Column \(CellAddress(column: columnIndex, row: 0).columnLetter)"
    }
}

/// Renames a sheet.
public class RenameSheetCommand: SpreadsheetCommand {
    public let sheetIndex: Int
    public let oldName: String
    public let newName: String

    public init(sheetIndex: Int, oldName: String, newName: String) {
        self.sheetIndex = sheetIndex
        self.oldName = oldName
        self.newName = newName
    }

    public func execute(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        workbook.sheets[sheetIndex].name = newName
    }

    public func undo(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        workbook.sheets[sheetIndex].name = oldName
    }

    public var description: String { "Rename Sheet" }
}

/// Moves a sheet to a final zero-based index while preserving active sheet identity.
public class MoveSheetCommand: SpreadsheetCommand {
    public let fromIndex: Int
    public let toIndex: Int
    private var movedSheetID: UUID?
    private var activeSheetID: UUID?

    public init(fromIndex: Int, toIndex: Int) {
        self.fromIndex = fromIndex
        self.toIndex = toIndex
    }

    public func execute(on workbook: Workbook) {
        guard fromIndex >= 0, fromIndex < workbook.sheets.count else { return }
        let sheetID = workbook.sheets[fromIndex].id
        movedSheetID = sheetID
        activeSheetID = workbook.sheets.indices.contains(workbook.activeSheetIndex)
            ? workbook.sheets[workbook.activeSheetIndex].id
            : nil
        move(sheetID: sheetID, to: toIndex, in: workbook)
        restoreActiveSheet(in: workbook)
    }

    public func undo(on workbook: Workbook) {
        guard let sheetID = movedSheetID else { return }
        move(sheetID: sheetID, to: fromIndex, in: workbook)
        restoreActiveSheet(in: workbook)
    }

    private func move(sheetID: UUID, to targetIndex: Int, in workbook: Workbook) {
        guard let current = workbook.sheets.firstIndex(where: { $0.id == sheetID }) else { return }
        let sheet = workbook.sheets.remove(at: current)
        let boundedTarget = max(0, min(targetIndex, workbook.sheets.count))
        workbook.sheets.insert(sheet, at: boundedTarget)
    }

    private func restoreActiveSheet(in workbook: Workbook) {
        guard let activeSheetID,
              let index = workbook.sheets.firstIndex(where: { $0.id == activeSheetID }) else { return }
        workbook.activeSheetIndex = index
    }

    public var description: String { "Move Sheet" }
}

/// Adds a sheet at the end of the workbook.
public class AddSheetCommand: SpreadsheetCommand {
    public let name: String?
    private var insertedSheet: Worksheet?
    private var insertedIndex: Int?

    public init(name: String? = nil) {
        self.name = name
    }

    public func execute(on workbook: Workbook) {
        let sheet = insertedSheet ?? Worksheet(name: name ?? nextSheetName(in: workbook))
        insertedSheet = sheet
        insertedIndex = workbook.sheets.count
        workbook.sheets.append(sheet)
    }

    public func undo(on workbook: Workbook) {
        guard let sheetID = insertedSheet?.id,
              let index = workbook.sheets.firstIndex(where: { $0.id == sheetID }),
              workbook.sheets.count > 1 else { return }
        workbook.sheets.remove(at: index)
        if workbook.activeSheetIndex >= workbook.sheets.count {
            workbook.activeSheetIndex = workbook.sheets.count - 1
        }
    }

    private func nextSheetName(in workbook: Workbook) -> String {
        var index = workbook.sheets.count + 1
        while workbook.sheets.contains(where: { $0.name == "Sheet\(index)" }) {
            index += 1
        }
        return "Sheet\(index)"
    }

    public var description: String { "Add Sheet" }
}

/// Deletes a sheet and keeps a deep copy for undo.
public class DeleteSheetCommand: SpreadsheetCommand {
    public let sheetIndex: Int
    private var deletedSheet: Worksheet?
    private var activeSheetID: UUID?

    public init(sheetIndex: Int) {
        self.sheetIndex = sheetIndex
    }

    public func execute(on workbook: Workbook) {
        guard workbook.sheets.count > 1,
              sheetIndex >= 0,
              sheetIndex < workbook.sheets.count else { return }
        activeSheetID = workbook.sheets.indices.contains(workbook.activeSheetIndex)
            ? workbook.sheets[workbook.activeSheetIndex].id
            : nil
        deletedSheet = workbook.sheets[sheetIndex].copy()
        workbook.sheets.remove(at: sheetIndex)
        if workbook.activeSheetIndex >= workbook.sheets.count {
            workbook.activeSheetIndex = workbook.sheets.count - 1
        }
    }

    public func undo(on workbook: Workbook) {
        guard let deletedSheet else { return }
        let target = max(0, min(sheetIndex, workbook.sheets.count))
        workbook.sheets.insert(deletedSheet.copy(), at: target)
        if let activeSheetID,
           let index = workbook.sheets.firstIndex(where: { $0.id == activeSheetID }) {
            workbook.activeSheetIndex = index
        }
    }

    public var description: String { "Delete Sheet" }
}

/// Duplicates a sheet next to its source.
public class DuplicateSheetCommand: SpreadsheetCommand {
    public let sourceIndex: Int
    private var duplicateSheet: Worksheet?

    public init(sourceIndex: Int) {
        self.sourceIndex = sourceIndex
    }

    public func execute(on workbook: Workbook) {
        guard sourceIndex >= 0, sourceIndex < workbook.sheets.count else { return }
        let source = workbook.sheets[sourceIndex]
        let duplicate = duplicateSheet ?? source.copy(name: "\(source.name) Copy", id: UUID())
        duplicateSheet = duplicate
        workbook.sheets.insert(duplicate, at: sourceIndex + 1)
    }

    public func undo(on workbook: Workbook) {
        guard let duplicateID = duplicateSheet?.id,
              let index = workbook.sheets.firstIndex(where: { $0.id == duplicateID }),
              workbook.sheets.count > 1 else { return }
        workbook.sheets.remove(at: index)
        if workbook.activeSheetIndex >= workbook.sheets.count {
            workbook.activeSheetIndex = workbook.sheets.count - 1
        }
    }

    public var description: String { "Duplicate Sheet" }
}

/// Inserts a row at the given index in a worksheet
public class InsertRowCommand: SpreadsheetCommand {
    public let sheetIndex: Int
    public let rowIndex: Int

    public init(sheetIndex: Int, rowIndex: Int) {
        self.sheetIndex = sheetIndex
        self.rowIndex = rowIndex
    }

    public func execute(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        workbook.sheets[sheetIndex].insertRow(at: rowIndex)
    }

    public func undo(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        workbook.sheets[sheetIndex].deleteRow(at: rowIndex)
    }

    public var description: String { "Insert Row \(rowIndex + 1)" }
}

/// Deletes a row at the given index in a worksheet
public class DeleteRowCommand: SpreadsheetCommand {
    public let sheetIndex: Int
    public let rowIndex: Int
    /// Saved cells and row height for undo
    private var savedCells: [CellAddress: Cell] = [:]
    private var savedRowHeight: Double?

    public init(sheetIndex: Int, rowIndex: Int) {
        self.sheetIndex = sheetIndex
        self.rowIndex = rowIndex
    }

    public func execute(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        let sheet = workbook.sheets[sheetIndex]

        // Save the cells in this row before deleting
        savedCells = [:]
        for (addr, cell) in sheet.cells where addr.row == rowIndex {
            savedCells[addr] = cell.copy()
        }
        savedRowHeight = sheet.rowHeights[rowIndex]

        sheet.deleteRow(at: rowIndex)
    }

    public func undo(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        let sheet = workbook.sheets[sheetIndex]

        // Re-insert the row
        sheet.insertRow(at: rowIndex)

        // Restore saved cells
        for (addr, cell) in savedCells {
            sheet.cells[addr] = cell
        }

        // Restore row height
        if let height = savedRowHeight {
            sheet.rowHeights[rowIndex] = height
        }
    }

    public var description: String { "Delete Row \(rowIndex + 1)" }
}

/// Inserts a column at the given index in a worksheet
public class InsertColumnCommand: SpreadsheetCommand {
    public let sheetIndex: Int
    public let columnIndex: Int

    public init(sheetIndex: Int, columnIndex: Int) {
        self.sheetIndex = sheetIndex
        self.columnIndex = columnIndex
    }

    public func execute(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        workbook.sheets[sheetIndex].insertColumn(at: columnIndex)
    }

    public func undo(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        workbook.sheets[sheetIndex].deleteColumn(at: columnIndex)
    }

    public var description: String {
        let letter = CellAddress(column: columnIndex, row: 0).columnLetter
        return "Insert Column \(letter)"
    }
}

/// Deletes a column at the given index in a worksheet
public class DeleteColumnCommand: SpreadsheetCommand {
    public let sheetIndex: Int
    public let columnIndex: Int
    /// Saved cells and column width for undo
    private var savedCells: [CellAddress: Cell] = [:]
    private var savedColumnWidth: Double?

    public init(sheetIndex: Int, columnIndex: Int) {
        self.sheetIndex = sheetIndex
        self.columnIndex = columnIndex
    }

    public func execute(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        let sheet = workbook.sheets[sheetIndex]

        // Save the cells in this column before deleting
        savedCells = [:]
        for (addr, cell) in sheet.cells where addr.column == columnIndex {
            savedCells[addr] = cell.copy()
        }
        savedColumnWidth = sheet.columnWidths[columnIndex]

        sheet.deleteColumn(at: columnIndex)
    }

    public func undo(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        let sheet = workbook.sheets[sheetIndex]

        // Re-insert the column
        sheet.insertColumn(at: columnIndex)

        // Restore saved cells
        for (addr, cell) in savedCells {
            sheet.cells[addr] = cell
        }

        // Restore column width
        if let width = savedColumnWidth {
            sheet.columnWidths[columnIndex] = width
        }
    }

    public var description: String {
        let letter = CellAddress(column: columnIndex, row: 0).columnLetter
        return "Delete Column \(letter)"
    }
}

/// A compound command that wraps multiple commands as a single undo group
public class CompoundCommand: SpreadsheetCommand {
    public let commands: [SpreadsheetCommand]
    public let label: String

    public init(commands: [SpreadsheetCommand], label: String = "Batch Edit") {
        self.commands = commands
        self.label = label
    }

    public func execute(on workbook: Workbook) {
        for command in commands {
            command.execute(on: workbook)
        }
    }

    public func undo(on workbook: Workbook) {
        for command in commands.reversed() {
            command.undo(on: workbook)
        }
    }

    public var description: String { label }
}

/// Clears all cells in a given range
public class ClearRangeCommand: SpreadsheetCommand {
    public let sheetIndex: Int
    public let range: CellRange
    /// Saved cells for undo
    private var savedCells: [CellAddress: Cell] = [:]

    public init(sheetIndex: Int, range: CellRange) {
        self.sheetIndex = sheetIndex
        self.range = range
    }

    public func execute(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        let sheet = workbook.sheets[sheetIndex]

        // Save existing cells in range before clearing
        savedCells = [:]
        for addr in range.allAddresses {
            if let cell = sheet.cells[addr] {
                savedCells[addr] = cell.copy()
            }
        }

        sheet.clearRange(range)
    }

    public func undo(on workbook: Workbook) {
        guard sheetIndex < workbook.sheets.count else { return }
        let sheet = workbook.sheets[sheetIndex]

        // Restore saved cells
        for (addr, cell) in savedCells {
            sheet.cells[addr] = cell
        }
    }

    public var description: String { "Clear \(range.displayString)" }
}
