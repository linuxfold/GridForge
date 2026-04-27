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

/// A formula-level cell reference.
///
/// `CellAddress` intentionally stays as a compact grid coordinate. Formula code
/// uses `CellReference` so absolute markers and future sheet-qualified
/// references are not lost during parsing.
public struct CellReference: Equatable, Hashable, Sendable {
    public var sheetID: UUID?
    public var column: Int
    public var row: Int
    public var columnAbsolute: Bool
    public var rowAbsolute: Bool

    public init(
        sheetID: UUID? = nil,
        column: Int,
        row: Int,
        columnAbsolute: Bool = false,
        rowAbsolute: Bool = false
    ) {
        self.sheetID = sheetID
        self.column = max(0, column)
        self.row = max(0, row)
        self.columnAbsolute = columnAbsolute
        self.rowAbsolute = rowAbsolute
    }

    public init(address: CellAddress, sheetID: UUID? = nil) {
        self.init(sheetID: sheetID, column: address.column, row: address.row)
    }

    public var address: CellAddress {
        CellAddress(column: column, row: row)
    }

    public var columnLetter: String {
        address.columnLetter
    }

    public var displayString: String {
        "\(columnAbsolute ? "$" : "")\(columnLetter)\(rowAbsolute ? "$" : "")\(row + 1)"
    }

    /// Parse a reference like "A1", "$A1", "A$1", or "$A$1".
    public static func parse(_ string: String, sheetID: UUID? = nil) -> CellReference? {
        let chars = Array(string.uppercased().trimmingCharacters(in: .whitespaces))
        guard !chars.isEmpty else { return nil }

        var index = 0
        var columnAbsolute = false
        var rowAbsolute = false

        if index < chars.count, chars[index] == "$" {
            columnAbsolute = true
            index += 1
        }

        let columnStart = index
        while index < chars.count, chars[index].isLetter, chars[index].isASCII {
            index += 1
        }
        guard index > columnStart else { return nil }

        if index < chars.count, chars[index] == "$" {
            rowAbsolute = true
            index += 1
        }

        let rowStart = index
        while index < chars.count, chars[index].isNumber {
            index += 1
        }
        guard index > rowStart, index == chars.count else { return nil }

        let columnLetters = String(chars[columnStart..<rowStart]).replacingOccurrences(of: "$", with: "")
        let rowString = String(chars[rowStart..<index])
        guard let rowNumber = Int(rowString), rowNumber > 0 else { return nil }

        return CellReference(
            sheetID: sheetID,
            column: CellAddress.columnLetterToIndex(columnLetters),
            row: rowNumber - 1,
            columnAbsolute: columnAbsolute,
            rowAbsolute: rowAbsolute
        )
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

/// A rectangular formula range that preserves reference metadata.
public struct CellRangeReference: Equatable, Hashable, Sendable {
    public let start: CellReference
    public let end: CellReference

    public init(start: CellReference, end: CellReference) {
        let normalized = CellRange(start: start.address, end: end.address)
        self.start = CellReference(
            sheetID: start.sheetID,
            column: normalized.start.column,
            row: normalized.start.row,
            columnAbsolute: start.columnAbsolute,
            rowAbsolute: start.rowAbsolute
        )
        self.end = CellReference(
            sheetID: end.sheetID ?? start.sheetID,
            column: normalized.end.column,
            row: normalized.end.row,
            columnAbsolute: end.columnAbsolute,
            rowAbsolute: end.rowAbsolute
        )
    }

    public var addressRange: CellRange {
        CellRange(start: start.address, end: end.address)
    }

    public var allReferences: [CellReference] {
        let range = addressRange
        var result: [CellReference] = []
        result.reserveCapacity(range.cellCount)
        for row in range.start.row...range.end.row {
            for column in range.start.column...range.end.column {
                result.append(CellReference(sheetID: start.sheetID, column: column, row: row))
            }
        }
        return result
    }
}
