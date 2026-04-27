import Foundation

// MARK: - AST Nodes

/// Operators for binary expressions.
public enum BinaryOperator: Equatable, Sendable {
    case add
    case subtract
    case multiply
    case divide
    case power
    case concatenate
    case equal
    case notEqual
    case lessThan
    case greaterThan
    case lessEqual
    case greaterEqual
}

/// Operators for unary expressions.
public enum UnaryOperator: Equatable, Sendable {
    case negate
    case percent
}

/// Abstract syntax tree node for formula expressions.
public indirect enum ASTNode: Equatable, Sendable {
    case number(Double)
    case string(String)
    case boolean(Bool)
    case cellReference(CellReference)
    case range(CellReference, CellReference)
    case binaryOp(BinaryOperator, ASTNode, ASTNode)
    case unaryOp(UnaryOperator, ASTNode)
    case functionCall(String, [ASTNode])
    case error(CellError)
}

// MARK: - Parser

/// Recursive descent parser that converts a token stream into an AST.
///
/// Precedence (low to high):
/// 1. Comparison: =, <>, <, >, <=, >=
/// 2. Addition/Subtraction: +, -
/// 3. Multiplication/Division: *, /
/// 4. Unary: - (negate)
/// 5. Power: ^
/// 6. Postfix: %
/// 7. Primary: numbers, strings, booleans, cell references, ranges, function calls, parenthesized expressions
public final class FormulaParser {

    private let tokens: [Token]
    private var current: Int = 0

    public init(tokens: [Token]) {
        self.tokens = tokens
    }

    /// Parse the token stream and return the root AST node.
    public func parse() throws -> ASTNode {
        let node = try parseComparison()
        guard currentToken.type == .eof else {
            throw ParserError.unexpectedToken(currentToken)
        }
        return node
    }

    // MARK: - Precedence Levels

    /// Comparison: =, <>, <, >, <=, >=
    private func parseComparison() throws -> ASTNode {
        var left = try parseAddition()

        while true {
            let op: BinaryOperator
            switch currentToken.type {
            case .equal:        op = .equal
            case .notEqual:     op = .notEqual
            case .lessThan:     op = .lessThan
            case .greaterThan:  op = .greaterThan
            case .lessEqual:    op = .lessEqual
            case .greaterEqual: op = .greaterEqual
            default: return left
            }
            advance()
            let right = try parseAddition()
            left = .binaryOp(op, left, right)
        }
    }

    /// Addition / Subtraction / Concatenation: +, -, &
    private func parseAddition() throws -> ASTNode {
        var left = try parseMultiplication()

        while true {
            let op: BinaryOperator
            switch currentToken.type {
            case .plus:      op = .add
            case .minus:     op = .subtract
            case .ampersand: op = .concatenate
            default: return left
            }
            advance()
            let right = try parseMultiplication()
            left = .binaryOp(op, left, right)
        }
    }

    /// Multiplication / Division: *, /
    private func parseMultiplication() throws -> ASTNode {
        var left = try parseUnary()

        while true {
            let op: BinaryOperator
            switch currentToken.type {
            case .multiply: op = .multiply
            case .divide:   op = .divide
            default: return left
            }
            advance()
            let right = try parseUnary()
            left = .binaryOp(op, left, right)
        }
    }

    /// Unary: -expr
    private func parseUnary() throws -> ASTNode {
        if currentToken.type == .minus {
            advance()
            let operand = try parseUnary()
            return .unaryOp(.negate, operand)
        }
        if currentToken.type == .plus {
            advance()
            return try parseUnary()
        }
        return try parsePower()
    }

    /// Power: base ^ exponent (right-associative)
    private func parsePower() throws -> ASTNode {
        var base = try parsePostfix()

        if currentToken.type == .power {
            advance()
            let exponent = try parseUnary()  // right-associative: recurse to unary
            base = .binaryOp(.power, base, exponent)
        }

        return base
    }

    /// Postfix: expr%
    private func parsePostfix() throws -> ASTNode {
        var node = try parsePrimary()

        while currentToken.type == .percent {
            advance()
            node = .unaryOp(.percent, node)
        }

        return node
    }

    /// Primary: literals, cell references, ranges, function calls, parenthesized expressions
    private func parsePrimary() throws -> ASTNode {
        let token = currentToken

        switch token.type {
        case .number(let value):
            advance()
            return .number(value)

        case .string(let value):
            advance()
            return .string(value)

        case .boolean(let value):
            advance()
            return .boolean(value)

        case .cellReference(let ref):
            advance()
            guard let reference = parseCellReferenceString(ref) else {
                return .error(.ref)
            }
            // Check for range operator ':'
            if currentToken.type == .colon {
                advance()
                guard case .cellReference(let endRef) = currentToken.type else {
                    throw ParserError.expectedCellReference(currentToken)
                }
                advance()
                guard let endAddress = parseCellReferenceString(endRef) else {
                    return .error(.ref)
                }
                return .range(reference, endAddress)
            }
            return .cellReference(reference)

        case .functionName(let name):
            advance()
            // Expect '('
            guard currentToken.type == .leftParen else {
                throw ParserError.expectedLeftParen(currentToken)
            }
            advance()

            // Parse arguments
            var arguments: [ASTNode] = []
            if currentToken.type != .rightParen {
                arguments.append(try parseComparison())
                while currentToken.type == .comma {
                    advance()
                    arguments.append(try parseComparison())
                }
            }

            guard currentToken.type == .rightParen else {
                throw ParserError.expectedRightParen(currentToken)
            }
            advance()

            return .functionCall(name, arguments)

        case .leftParen:
            advance()
            let expr = try parseComparison()
            guard currentToken.type == .rightParen else {
                throw ParserError.expectedRightParen(currentToken)
            }
            advance()
            return expr

        default:
            throw ParserError.unexpectedToken(token)
        }
    }

    // MARK: - Helpers

    private var currentToken: Token {
        guard current < tokens.count else {
            return Token(type: .eof, position: -1)
        }
        return tokens[current]
    }

    @discardableResult
    private func advance() -> Token {
        let token = currentToken
        current += 1
        return token
    }

    /// Parse a cell reference string into the richer formula reference model.
    private func parseCellReferenceString(_ ref: String) -> CellReference? {
        CellReference.parse(ref)
    }
}

// MARK: - Parser Errors

public enum ParserError: Error, Equatable, Sendable {
    case unexpectedToken(Token)
    case expectedCellReference(Token)
    case expectedLeftParen(Token)
    case expectedRightParen(Token)
}
