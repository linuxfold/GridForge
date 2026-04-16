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

// MARK: - Undo Manager

/// Command-pattern undo/redo manager for spreadsheet operations
public class SpreadsheetUndoManager {
    private var undoStack: [SpreadsheetCommand] = []
    private var redoStack: [SpreadsheetCommand] = []

    public init() {}

    /// Execute a command and push it onto the undo stack.
    /// Clears the redo stack (new action invalidates redo history).
    public func perform(_ command: SpreadsheetCommand, on workbook: Workbook) {
        command.execute(on: workbook)
        undoStack.append(command)
        redoStack.removeAll()
    }

    /// Undo the most recent command. Returns true if an undo was performed.
    @discardableResult
    public func undo(on workbook: Workbook) -> Bool {
        guard let command = undoStack.popLast() else { return false }
        command.undo(on: workbook)
        redoStack.append(command)
        return true
    }

    /// Redo the most recently undone command. Returns true if a redo was performed.
    @discardableResult
    public func redo(on workbook: Workbook) -> Bool {
        guard let command = redoStack.popLast() else { return false }
        command.execute(on: workbook)
        undoStack.append(command)
        return true
    }

    /// Whether there are commands available to undo
    public var canUndo: Bool { !undoStack.isEmpty }

    /// Whether there are commands available to redo
    public var canRedo: Bool { !redoStack.isEmpty }

    /// Clear all undo/redo history
    public func clear() {
        undoStack.removeAll()
        redoStack.removeAll()
    }

    /// The description of the next command to undo, if any
    public var undoDescription: String? {
        undoStack.last?.description
    }

    /// The description of the next command to redo, if any
    public var redoDescription: String? {
        redoStack.last?.description
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
