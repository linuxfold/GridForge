import Foundation

// MARK: - Token Types

/// The kind of token produced by the formula tokenizer.
public enum TokenType: Equatable, Sendable {
    // Literals
    case number(Double)
    case string(String)
    case boolean(Bool)

    // References and identifiers
    case cellReference(String)
    case functionName(String)

    // Arithmetic operators
    case plus
    case minus
    case multiply
    case divide
    case power
    case percent

    // String operator
    case ampersand

    // Comparison operators
    case equal
    case notEqual
    case lessThan
    case greaterThan
    case lessEqual
    case greaterEqual

    // Delimiters
    case leftParen
    case rightParen
    case comma
    case colon

    // Sentinel
    case eof
}

/// A single token with its type and position in the formula string.
public struct Token: Equatable, Sendable {
    public let type: TokenType
    public let position: Int

    public init(type: TokenType, position: Int) {
        self.type = type
        self.position = position
    }
}

// MARK: - Tokenizer

/// Converts a formula string (without the leading "=") into a sequence of tokens.
public final class Tokenizer {

    private let source: [Character]
    private var current: Int = 0
    private var tokens: [Token] = []

    /// Initialize with the formula text (without the leading "=").
    public init(formula: String) {
        self.source = Array(formula)
    }

    /// Tokenize the formula and return the token list (always ends with `.eof`).
    public func tokenize() throws -> [Token] {
        tokens = []
        current = 0

        while !isAtEnd {
            skipWhitespace()
            if isAtEnd { break }
            try scanToken()
        }

        tokens.append(Token(type: .eof, position: current))
        return tokens
    }

    // MARK: - Scanning

    private func scanToken() throws {
        let startPos = current
        let c = advance()

        switch c {
        case "+": tokens.append(Token(type: .plus, position: startPos))
        case "-": tokens.append(Token(type: .minus, position: startPos))
        case "*": tokens.append(Token(type: .multiply, position: startPos))
        case "/": tokens.append(Token(type: .divide, position: startPos))
        case "^": tokens.append(Token(type: .power, position: startPos))
        case "%": tokens.append(Token(type: .percent, position: startPos))
        case "&": tokens.append(Token(type: .ampersand, position: startPos))
        case "(": tokens.append(Token(type: .leftParen, position: startPos))
        case ")": tokens.append(Token(type: .rightParen, position: startPos))
        case ",": tokens.append(Token(type: .comma, position: startPos))
        case ":": tokens.append(Token(type: .colon, position: startPos))
        case "=": tokens.append(Token(type: .equal, position: startPos))
        case "<":
            if match("=") {
                tokens.append(Token(type: .lessEqual, position: startPos))
            } else if match(">") {
                tokens.append(Token(type: .notEqual, position: startPos))
            } else {
                tokens.append(Token(type: .lessThan, position: startPos))
            }
        case ">":
            if match("=") {
                tokens.append(Token(type: .greaterEqual, position: startPos))
            } else {
                tokens.append(Token(type: .greaterThan, position: startPos))
            }
        case "\"":
            try scanString(startPos: startPos)
        default:
            if c.isNumber || (c == "." && !isAtEnd && peek.isNumber) {
                scanNumber(startPos: startPos)
            } else if c == "$" || c.isLetter {
                scanIdentifierOrReference(startPos: startPos)
            } else {
                throw TokenizerError.unexpectedCharacter(c, position: startPos)
            }
        }
    }

    // MARK: - String Literals

    private func scanString(startPos: Int) throws {
        var value = ""
        while !isAtEnd && peek != "\"" {
            value.append(advance())
        }
        guard !isAtEnd else {
            throw TokenizerError.unterminatedString(position: startPos)
        }
        // consume closing quote
        _ = advance()
        tokens.append(Token(type: .string(value), position: startPos))
    }

    // MARK: - Numbers

    private func scanNumber(startPos: Int) {
        // We already consumed the first character; back up so we can read the whole number.
        current = startPos
        while !isAtEnd && peek.isNumber {
            _ = advance()
        }
        // Decimal part
        if !isAtEnd && peek == "." {
            _ = advance() // consume '.'
            while !isAtEnd && peek.isNumber {
                _ = advance()
            }
        }
        // Scientific notation (e.g. 1.5E+3)
        if !isAtEnd && (peek == "e" || peek == "E") {
            _ = advance()
            if !isAtEnd && (peek == "+" || peek == "-") {
                _ = advance()
            }
            while !isAtEnd && peek.isNumber {
                _ = advance()
            }
        }

        let text = String(source[startPos..<current])
        let value = Double(text) ?? 0
        tokens.append(Token(type: .number(value), position: startPos))
    }

    // MARK: - Identifiers, Cell References, Booleans

    private func scanIdentifierOrReference(startPos: Int) {
        // Back up so we can re-read from startPos.
        current = startPos

        // Consume optional leading '$' characters and letters/digits that form a reference or identifier.
        // Cell references can look like: A1, $A1, A$1, $A$1, AA12, $AA$12
        // Function names look like: SUM, AVERAGE, IF (followed by '(')
        // Booleans: TRUE, FALSE

        // Gather the full identifier-like token (letters, digits, $, _)
        while !isAtEnd && (peek.isLetter || peek.isNumber || peek == "$" || peek == "_") {
            _ = advance()
        }

        let text = String(source[startPos..<current])
        let upperText = text.uppercased()

        // Check for booleans
        if upperText == "TRUE" {
            tokens.append(Token(type: .boolean(true), position: startPos))
            return
        }
        if upperText == "FALSE" {
            tokens.append(Token(type: .boolean(false), position: startPos))
            return
        }

        // Check if this is a function name (followed by '(')
        skipWhitespace()
        if !isAtEnd && peek == "(" {
            tokens.append(Token(type: .functionName(upperText), position: startPos))
            return
        }

        // Check if this looks like a cell reference: optional $, one or more letters, optional $, one or more digits
        if isCellReference(text) {
            tokens.append(Token(type: .cellReference(text.uppercased()), position: startPos))
            return
        }

        // Unknown identifier — treat as a name error; emit as a function name
        // so the parser can produce a #NAME? error
        tokens.append(Token(type: .functionName(upperText), position: startPos))
    }

    /// Determines whether a string is a valid cell reference pattern.
    private func isCellReference(_ text: String) -> Bool {
        // Pattern: optional $, one or more ASCII letters, optional $, one or more digits
        let chars = Array(text.uppercased())
        var i = 0

        // Optional leading $
        if i < chars.count && chars[i] == "$" { i += 1 }

        // Must have at least one letter
        let letterStart = i
        while i < chars.count && chars[i].isLetter && chars[i].isASCII { i += 1 }
        guard i > letterStart else { return false }

        // Optional $ before row number
        if i < chars.count && chars[i] == "$" { i += 1 }

        // Must have at least one digit
        let digitStart = i
        while i < chars.count && chars[i].isNumber { i += 1 }
        guard i > digitStart else { return false }

        // Must have consumed entire string
        return i == chars.count
    }

    // MARK: - Character Helpers

    private var isAtEnd: Bool { current >= source.count }

    private var peek: Character {
        guard current < source.count else { return "\0" }
        return source[current]
    }

    @discardableResult
    private func advance() -> Character {
        let c = source[current]
        current += 1
        return c
    }

    private func match(_ expected: Character) -> Bool {
        guard !isAtEnd && peek == expected else { return false }
        current += 1
        return true
    }

    private func skipWhitespace() {
        while !isAtEnd && peek.isWhitespace {
            current += 1
        }
    }
}

// MARK: - Tokenizer Errors

public enum TokenizerError: Error, Equatable, Sendable {
    case unexpectedCharacter(Character, position: Int)
    case unterminatedString(position: Int)
}
