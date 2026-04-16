import Foundation

/// Direction for keyboard-based cell navigation
public enum Direction: Sendable {
    case up
    case down
    case left
    case right
}

/// Represents the current selection state in a spreadsheet view
public struct SelectionState: Equatable, Sendable {
    /// The currently active (focused) cell
    public var activeCell: CellAddress

    /// The selected rectangular range (may be a single cell)
    public var selectedRange: CellRange

    /// Whether the user is actively editing the cell content
    public var isEditing: Bool

    public init(activeCell: CellAddress = CellAddress(column: 0, row: 0)) {
        self.activeCell = activeCell
        self.selectedRange = CellRange(activeCell)
        self.isEditing = false
    }

    /// Select a single cell (collapses the range to one cell)
    public mutating func select(cell: CellAddress) {
        activeCell = cell
        selectedRange = CellRange(cell)
        isEditing = false
    }

    /// Extend the current selection from the active cell to the given cell
    /// The active cell remains the anchor; the range expands to include `to`
    public mutating func extendSelection(to target: CellAddress) {
        selectedRange = CellRange(start: activeCell, end: target)
    }

    /// Move the active cell in the given direction
    /// - Parameters:
    ///   - direction: The direction to move
    ///   - extend: If true, extends the selection rather than moving it
    ///   - maxColumn: Maximum column index (exclusive upper bound)
    ///   - maxRow: Maximum row index (exclusive upper bound)
    public mutating func moveActiveCell(
        direction: Direction,
        extend: Bool,
        maxColumn: Int,
        maxRow: Int
    ) {
        let current = extend ? selectionEdge(for: direction) : activeCell
        let newAddress: CellAddress

        switch direction {
        case .up:
            let newRow = max(0, current.row - 1)
            newAddress = CellAddress(column: current.column, row: newRow)
        case .down:
            let newRow = min(maxRow - 1, current.row + 1)
            newAddress = CellAddress(column: current.column, row: newRow)
        case .left:
            let newCol = max(0, current.column - 1)
            newAddress = CellAddress(column: newCol, row: current.row)
        case .right:
            let newCol = min(maxColumn - 1, current.column + 1)
            newAddress = CellAddress(column: newCol, row: current.row)
        }

        if extend {
            extendSelection(to: newAddress)
        } else {
            activeCell = newAddress
            selectedRange = CellRange(newAddress)
        }

        isEditing = false
    }

    /// Returns the edge of the selection range in the given direction
    /// Used to determine where to extend from when Shift+arrow is pressed
    private func selectionEdge(for direction: Direction) -> CellAddress {
        switch direction {
        case .up:
            // If active cell is at the bottom, use the top edge; otherwise use the bottom
            if activeCell.row == selectedRange.end.row {
                return selectedRange.start
            } else {
                return selectedRange.end
            }
        case .down:
            if activeCell.row == selectedRange.start.row {
                return selectedRange.end
            } else {
                return selectedRange.start
            }
        case .left:
            if activeCell.column == selectedRange.end.column {
                return CellAddress(column: selectedRange.start.column, row: activeCell.row)
            } else {
                return CellAddress(column: selectedRange.end.column, row: activeCell.row)
            }
        case .right:
            if activeCell.column == selectedRange.start.column {
                return CellAddress(column: selectedRange.end.column, row: activeCell.row)
            } else {
                return CellAddress(column: selectedRange.start.column, row: activeCell.row)
            }
        }
    }
}
