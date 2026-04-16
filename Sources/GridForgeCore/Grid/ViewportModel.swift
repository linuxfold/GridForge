import Foundation

/// Manages viewport calculations for a virtualized spreadsheet grid.
/// Determines which rows and columns are visible given scroll position and viewport size.
public struct ViewportModel {
    /// The range of row indices currently visible in the viewport
    public var visibleRows: Range<Int>

    /// The range of column indices currently visible in the viewport
    public var visibleColumns: Range<Int>

    /// Current horizontal scroll offset in points
    public var scrollOffsetX: Double

    /// Current vertical scroll offset in points
    public var scrollOffsetY: Double

    /// Width of the visible viewport area in points
    public var viewportWidth: Double

    /// Height of the visible viewport area in points
    public var viewportHeight: Double

    /// Width reserved for row number headers
    public static let rowHeaderWidth: Double = 50

    /// Height reserved for column letter headers
    public static let columnHeaderHeight: Double = 28

    public init() {
        self.visibleRows = 0..<1
        self.visibleColumns = 0..<1
        self.scrollOffsetX = 0
        self.scrollOffsetY = 0
        self.viewportWidth = 0
        self.viewportHeight = 0
    }

    /// Update the viewport model based on current scroll position and viewport size.
    /// Recalculates which rows and columns are visible.
    ///
    /// - Parameters:
    ///   - scrollOffset: Current scroll position (negative values for SwiftUI scroll views)
    ///   - viewportSize: Size of the visible area
    ///   - worksheet: The worksheet to calculate against (for custom row/column sizes)
    ///   - displayColumns: Total number of columns to consider
    ///   - displayRows: Total number of rows to consider
    public mutating func update(
        scrollOffset: CGPoint,
        viewportSize: CGSize,
        worksheet: Worksheet,
        displayColumns: Int,
        displayRows: Int
    ) {
        // Store raw values (scroll offsets may be negative in SwiftUI)
        scrollOffsetX = abs(Double(scrollOffset.x))
        scrollOffsetY = abs(Double(scrollOffset.y))
        viewportWidth = Double(viewportSize.width)
        viewportHeight = Double(viewportSize.height)

        // Available area after headers
        let contentWidth = max(0, viewportWidth - ViewportModel.rowHeaderWidth)
        let contentHeight = max(0, viewportHeight - ViewportModel.columnHeaderHeight)

        // Determine visible column range
        let firstCol = findFirstVisibleColumn(
            scrollX: scrollOffsetX,
            worksheet: worksheet,
            maxColumns: displayColumns
        )
        let lastCol = findLastVisibleColumn(
            firstColumn: firstCol,
            availableWidth: contentWidth,
            worksheet: worksheet,
            maxColumns: displayColumns
        )
        visibleColumns = firstCol..<(lastCol + 1)

        // Determine visible row range
        let firstRow = findFirstVisibleRow(
            scrollY: scrollOffsetY,
            worksheet: worksheet,
            maxRows: displayRows
        )
        let lastRow = findLastVisibleRow(
            firstRow: firstRow,
            availableHeight: contentHeight,
            worksheet: worksheet,
            maxRows: displayRows
        )
        visibleRows = firstRow..<(lastRow + 1)
    }

    /// Calculate total content width for the given number of columns
    public func totalContentWidth(worksheet: Worksheet, columns: Int) -> Double {
        var total: Double = 0
        for c in 0..<columns {
            total += worksheet.columnWidth(for: c)
        }
        return total
    }

    /// Calculate total content height for the given number of rows
    public func totalContentHeight(worksheet: Worksheet, rows: Int) -> Double {
        var total: Double = 0
        for r in 0..<rows {
            total += worksheet.rowHeight(for: r)
        }
        return total
    }

    // MARK: - Private Helpers

    /// Find the first column that is at least partially visible at the given scroll position
    private func findFirstVisibleColumn(
        scrollX: Double,
        worksheet: Worksheet,
        maxColumns: Int
    ) -> Int {
        var accumulatedWidth: Double = 0
        for col in 0..<maxColumns {
            let colWidth = worksheet.columnWidth(for: col)
            if accumulatedWidth + colWidth > scrollX {
                return col
            }
            accumulatedWidth += colWidth
        }
        return max(0, maxColumns - 1)
    }

    /// Find the last column that is at least partially visible given the first visible column
    private func findLastVisibleColumn(
        firstColumn: Int,
        availableWidth: Double,
        worksheet: Worksheet,
        maxColumns: Int
    ) -> Int {
        var usedWidth: Double = 0
        // Account for the portion of the first column that may be scrolled off
        // (already handled by starting from firstColumn)
        for col in firstColumn..<maxColumns {
            usedWidth += worksheet.columnWidth(for: col)
            if usedWidth >= availableWidth + worksheet.columnWidth(for: firstColumn) {
                return min(col, maxColumns - 1)
            }
        }
        return max(firstColumn, maxColumns - 1)
    }

    /// Find the first row that is at least partially visible at the given scroll position
    private func findFirstVisibleRow(
        scrollY: Double,
        worksheet: Worksheet,
        maxRows: Int
    ) -> Int {
        var accumulatedHeight: Double = 0
        for row in 0..<maxRows {
            let rowH = worksheet.rowHeight(for: row)
            if accumulatedHeight + rowH > scrollY {
                return row
            }
            accumulatedHeight += rowH
        }
        return max(0, maxRows - 1)
    }

    /// Find the last row that is at least partially visible given the first visible row
    private func findLastVisibleRow(
        firstRow: Int,
        availableHeight: Double,
        worksheet: Worksheet,
        maxRows: Int
    ) -> Int {
        var usedHeight: Double = 0
        for row in firstRow..<maxRows {
            usedHeight += worksheet.rowHeight(for: row)
            if usedHeight >= availableHeight + worksheet.rowHeight(for: firstRow) {
                return min(row, maxRows - 1)
            }
        }
        return max(firstRow, maxRows - 1)
    }
}
