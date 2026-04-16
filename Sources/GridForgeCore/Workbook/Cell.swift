import Foundation

/// A single cell in a spreadsheet worksheet
public final class Cell {
    /// Exactly what the user typed or what was imported
    public var rawInput: String
    /// Resolved/computed value (formula result or parsed input)
    public var value: CellValue
    /// Visual formatting
    public var formatting: CellFormatting

    public init(rawInput: String = "", value: CellValue = .empty, formatting: CellFormatting = CellFormatting()) {
        self.rawInput = rawInput
        self.value = value
        self.formatting = formatting
    }

    /// Whether the raw input is a formula (starts with "=")
    public var isFormula: Bool { rawInput.hasPrefix("=") }

    /// The formula expression without the leading "="
    public var formulaExpression: String? {
        guard isFormula else { return nil }
        return String(rawInput.dropFirst())
    }

    /// What to show in the cell
    public var displayString: String { value.displayString }

    /// What to show in the formula bar
    public var editString: String {
        rawInput.isEmpty ? value.displayString : rawInput
    }

    /// Whether this cell has no content
    public var isEmpty: Bool { rawInput.isEmpty && value == .empty }

    /// Deep copy
    public func copy() -> Cell {
        Cell(rawInput: rawInput, value: value, formatting: formatting)
    }
}

/// Cell visual formatting properties
public struct CellFormatting: Equatable, Sendable {
    public var bold: Bool
    public var italic: Bool
    public var underline: Bool
    public var alignment: HorizontalAlignment
    public var numberFormat: String?
    public var fontSize: Double
    public var fontName: String?
    public var textColor: CellColor?
    public var backgroundColor: CellColor?

    public init(
        bold: Bool = false,
        italic: Bool = false,
        underline: Bool = false,
        alignment: HorizontalAlignment = .general,
        numberFormat: String? = nil,
        fontSize: Double = 13,
        fontName: String? = nil,
        textColor: CellColor? = nil,
        backgroundColor: CellColor? = nil
    ) {
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.alignment = alignment
        self.numberFormat = numberFormat
        self.fontSize = fontSize
        self.fontName = fontName
        self.textColor = textColor
        self.backgroundColor = backgroundColor
    }
}

public enum HorizontalAlignment: String, Equatable, Sendable {
    case general  // numbers right-aligned, text left-aligned
    case left
    case center
    case right
}

public struct CellColor: Equatable, Hashable, Sendable {
    public let red: Double
    public let green: Double
    public let blue: Double
    public let alpha: Double

    public init(red: Double, green: Double, blue: Double, alpha: Double = 1.0) {
        self.red = red
        self.green = green
        self.blue = blue
        self.alpha = alpha
    }

    public static let black = CellColor(red: 0, green: 0, blue: 0)
    public static let white = CellColor(red: 1, green: 1, blue: 1)
}
