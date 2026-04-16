import Foundation

/// A single sheet within a workbook
public final class Worksheet: Identifiable {
    public let id: UUID
    public var name: String
    public var cells: [CellAddress: Cell]
    public var columnWidths: [Int: Double]
    public var rowHeights: [Int: Double]

    public static let defaultColumnWidth: Double = 100
    public static let defaultRowHeight: Double = 24
    public static let maxColumns: Int = 702   // A–ZZ
    public static let maxRows: Int = 100_000

    // Cached cumulative offsets for O(1) position lookups
    private var _columnOffsets: [Double]?
    private var _rowOffsets: [Double]?
    private var _cachedColumnCount: Int = 0
    private var _cachedRowCount: Int = 0

    public init(name: String, id: UUID = UUID()) {
        self.id = id
        self.name = name
        self.cells = [:]
        self.columnWidths = [:]
        self.rowHeights = [:]
    }

    // MARK: - Cell Access

    public func cell(at address: CellAddress) -> Cell? {
        cells[address]
    }

    public func cellValue(at address: CellAddress) -> CellValue {
        cells[address]?.value ?? .empty
    }

    @discardableResult
    public func setRawInput(at address: CellAddress, rawInput: String) -> Cell {
        if rawInput.isEmpty {
            cells.removeValue(forKey: address)
            return Cell()
        }
        let cell: Cell
        if let existing = cells[address] {
            existing.rawInput = rawInput
            cell = existing
        } else {
            cell = Cell(rawInput: rawInput)
            cells[address] = cell
        }
        if !cell.isFormula {
            cell.value = Worksheet.parseRawInput(rawInput)
        }
        return cell
    }

    public func clearCell(at address: CellAddress) {
        cells.removeValue(forKey: address)
    }

    public func clearRange(_ range: CellRange) {
        for addr in range.allAddresses {
            cells.removeValue(forKey: addr)
        }
    }

    // MARK: - Dimensions

    public func columnWidth(for column: Int) -> Double {
        columnWidths[column] ?? Worksheet.defaultColumnWidth
    }

    public func rowHeight(for row: Int) -> Double {
        rowHeights[row] ?? Worksheet.defaultRowHeight
    }

    public func setColumnWidth(_ width: Double, for column: Int) {
        columnWidths[column] = max(20, min(500, width))
        invalidateColumnCache()
    }

    public func setRowHeight(_ height: Double, for row: Int) {
        rowHeights[row] = max(14, min(200, height))
        invalidateRowCache()
    }

    // MARK: - Cached Offset Computation (O(1) lookups after build)

    /// Invalidate caches when widths/heights change
    public func invalidateColumnCache() { _columnOffsets = nil }
    public func invalidateRowCache() { _rowOffsets = nil }
    public func invalidateLayoutCaches() { _columnOffsets = nil; _rowOffsets = nil }

    /// Build or return cached column offset array.
    /// offsets[i] = x position of column i (cumulative sum of widths before it).
    public func columnOffsets(count: Int) -> [Double] {
        if let cached = _columnOffsets, _cachedColumnCount == count { return cached }
        var offsets = [Double]()
        offsets.reserveCapacity(count + 1)
        var x: Double = 0
        for c in 0..<count {
            offsets.append(x)
            x += columnWidth(for: c)
        }
        offsets.append(x) // trailing edge of last column
        _columnOffsets = offsets
        _cachedColumnCount = count
        return offsets
    }

    /// Build or return cached row offset array.
    public func rowOffsets(count: Int) -> [Double] {
        if let cached = _rowOffsets, _cachedRowCount == count { return cached }
        var offsets = [Double]()
        offsets.reserveCapacity(count + 1)
        var y: Double = 0
        for r in 0..<count {
            offsets.append(y)
            y += rowHeight(for: r)
        }
        offsets.append(y)
        _rowOffsets = offsets
        _cachedRowCount = count
        return offsets
    }

    /// X position of column (uses cache)
    public func xOffset(for column: Int, totalColumns: Int) -> Double {
        let offsets = columnOffsets(count: totalColumns)
        guard column < offsets.count else { return offsets.last ?? 0 }
        return offsets[column]
    }

    /// Y position of row (uses cache)
    public func yOffset(for row: Int, totalRows: Int) -> Double {
        let offsets = rowOffsets(count: totalRows)
        guard row < offsets.count else { return offsets.last ?? 0 }
        return offsets[row]
    }

    /// Total content width for N columns
    public func totalWidth(columns: Int) -> Double {
        let offsets = columnOffsets(count: columns)
        return offsets.last ?? 0
    }

    /// Total content height for N rows
    public func totalHeight(rows: Int) -> Double {
        let offsets = rowOffsets(count: rows)
        return offsets.last ?? 0
    }

    /// Binary search: find column at x position
    public func columnAt(x: Double, totalColumns: Int) -> Int {
        let offsets = columnOffsets(count: totalColumns)
        var lo = 0, hi = totalColumns - 1
        while lo <= hi {
            let mid = (lo + hi) / 2
            if offsets[mid + 1] <= x {
                lo = mid + 1
            } else if offsets[mid] > x {
                hi = mid - 1
            } else {
                return mid
            }
        }
        return max(0, min(totalColumns - 1, lo))
    }

    /// Binary search: find row at y position
    public func rowAt(y: Double, totalRows: Int) -> Int {
        let offsets = rowOffsets(count: totalRows)
        var lo = 0, hi = totalRows - 1
        while lo <= hi {
            let mid = (lo + hi) / 2
            if offsets[mid + 1] <= y {
                lo = mid + 1
            } else if offsets[mid] > y {
                hi = mid - 1
            } else {
                return mid
            }
        }
        return max(0, min(totalRows - 1, lo))
    }

    // Legacy compat (uncached, for use when totalColumns unknown)
    public func xOffset(for column: Int) -> Double {
        var x: Double = 0
        for c in 0..<column { x += columnWidth(for: c) }
        return x
    }

    public func yOffset(for row: Int) -> Double {
        var y: Double = 0
        for r in 0..<row { y += rowHeight(for: r) }
        return y
    }

    /// Bounding rectangle of used area
    public var usedRange: CellRange? {
        guard !cells.isEmpty else { return nil }
        var minCol = Int.max, maxCol = 0
        var minRow = Int.max, maxRow = 0
        for addr in cells.keys {
            minCol = min(minCol, addr.column)
            maxCol = max(maxCol, addr.column)
            minRow = min(minRow, addr.row)
            maxRow = max(maxRow, addr.row)
        }
        return CellRange(
            start: CellAddress(column: minCol, row: minRow),
            end: CellAddress(column: maxCol, row: maxRow)
        )
    }

    // MARK: - Row / Column Operations

    public func insertRow(at rowIndex: Int) {
        var newCells: [CellAddress: Cell] = [:]
        for (addr, cell) in cells {
            if addr.row >= rowIndex {
                newCells[CellAddress(column: addr.column, row: addr.row + 1)] = cell
            } else {
                newCells[addr] = cell
            }
        }
        cells = newCells
        var h: [Int: Double] = [:]
        for (row, height) in rowHeights {
            h[row >= rowIndex ? row + 1 : row] = height
        }
        rowHeights = h
        invalidateRowCache()
    }

    public func deleteRow(at rowIndex: Int) {
        var newCells: [CellAddress: Cell] = [:]
        for (addr, cell) in cells {
            if addr.row == rowIndex { continue }
            if addr.row > rowIndex {
                newCells[CellAddress(column: addr.column, row: addr.row - 1)] = cell
            } else {
                newCells[addr] = cell
            }
        }
        cells = newCells
        invalidateRowCache()
    }

    public func insertColumn(at colIndex: Int) {
        var newCells: [CellAddress: Cell] = [:]
        for (addr, cell) in cells {
            if addr.column >= colIndex {
                newCells[CellAddress(column: addr.column + 1, row: addr.row)] = cell
            } else {
                newCells[addr] = cell
            }
        }
        cells = newCells
        var w: [Int: Double] = [:]
        for (col, width) in columnWidths {
            w[col >= colIndex ? col + 1 : col] = width
        }
        columnWidths = w
        invalidateColumnCache()
    }

    public func deleteColumn(at colIndex: Int) {
        var newCells: [CellAddress: Cell] = [:]
        for (addr, cell) in cells {
            if addr.column == colIndex { continue }
            if addr.column > colIndex {
                newCells[CellAddress(column: addr.column - 1, row: addr.row)] = cell
            } else {
                newCells[addr] = cell
            }
        }
        cells = newCells
        invalidateColumnCache()
    }

    // MARK: - Input Parsing

    public static func parseRawInput(_ input: String) -> CellValue {
        let trimmed = input.trimmingCharacters(in: .whitespaces)
        if trimmed.isEmpty { return .empty }
        if let n = Double(trimmed) { return .number(n) }
        let upper = trimmed.uppercased()
        if upper == "TRUE" { return .boolean(true) }
        if upper == "FALSE" { return .boolean(false) }
        return .string(trimmed)
    }
}
