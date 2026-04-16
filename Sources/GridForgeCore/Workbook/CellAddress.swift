import Foundation

/// Zero-based cell coordinate in a spreadsheet
public struct CellAddress: Equatable, Hashable, Sendable, Comparable {
    public let column: Int
    public let row: Int

    public init(column: Int, row: Int) {
        self.column = max(0, column)
        self.row = max(0, row)
    }

    /// Column letter(s): 0→A, 1→B, ..., 25→Z, 26→AA
    public var columnLetter: String {
        var result = ""
        var n = column
        repeat {
            result = String(UnicodeScalar(65 + n % 26)!) + result
            n = n / 26 - 1
        } while n >= 0
        return result
    }

    /// Display string like "A1", "B12"
    public var displayString: String {
        "\(columnLetter)\(row + 1)"
    }

    /// Parse "A1" → CellAddress(column: 0, row: 0)
    public static func parse(_ string: String) -> CellAddress? {
        let s = string.uppercased().trimmingCharacters(in: .whitespaces)
        guard !s.isEmpty else { return nil }

        var colPart = ""
        var rowPart = ""
        for char in s {
            if char.isLetter && char.isASCII {
                if !rowPart.isEmpty { return nil }
                colPart.append(char)
            } else if char.isNumber {
                rowPart.append(char)
            } else {
                return nil
            }
        }

        guard !colPart.isEmpty, !rowPart.isEmpty,
              let rowNum = Int(rowPart), rowNum > 0 else { return nil }

        var col = 0
        for char in colPart {
            col = col * 26 + Int(char.asciiValue! - 65) + 1
        }
        col -= 1

        return CellAddress(column: col, row: rowNum - 1)
    }

    /// Convert column letters to zero-based index
    public static func columnLetterToIndex(_ letters: String) -> Int {
        var col = 0
        for char in letters.uppercased() {
            guard let ascii = char.asciiValue else { continue }
            col = col * 26 + Int(ascii - 65) + 1
        }
        return col - 1
    }

    public static func < (lhs: CellAddress, rhs: CellAddress) -> Bool {
        if lhs.row != rhs.row { return lhs.row < rhs.row }
        return lhs.column < rhs.column
    }
}

/// A rectangular range of cells
public struct CellRange: Equatable, Hashable, Sendable {
    public let start: CellAddress
    public let end: CellAddress

    /// Normalizes so start ≤ end
    public init(start: CellAddress, end: CellAddress) {
        self.start = CellAddress(
            column: min(start.column, end.column),
            row: min(start.row, end.row)
        )
        self.end = CellAddress(
            column: max(start.column, end.column),
            row: max(start.row, end.row)
        )
    }

    /// Single-cell range
    public init(_ address: CellAddress) {
        self.start = address
        self.end = address
    }

    public var isSingleCell: Bool { start == end }
    public var rowCount: Int { end.row - start.row + 1 }
    public var columnCount: Int { end.column - start.column + 1 }
    public var cellCount: Int { rowCount * columnCount }

    public func contains(_ address: CellAddress) -> Bool {
        address.row >= start.row && address.row <= end.row &&
        address.column >= start.column && address.column <= end.column
    }

    public var allAddresses: [CellAddress] {
        var result: [CellAddress] = []
        result.reserveCapacity(cellCount)
        for r in start.row...end.row {
            for c in start.column...end.column {
                result.append(CellAddress(column: c, row: r))
            }
        }
        return result
    }

    public var displayString: String {
        isSingleCell ? start.displayString : "\(start.displayString):\(end.displayString)"
    }
}
