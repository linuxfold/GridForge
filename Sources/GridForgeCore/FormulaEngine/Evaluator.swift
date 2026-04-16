import Foundation

// MARK: - Cell Value Provider Protocol

/// Abstraction for looking up cell values, allowing the evaluator to work
/// without depending directly on Worksheet.
public protocol CellValueProvider {
    func cellValue(at address: CellAddress) -> CellValue
}

// MARK: - Formula Evaluator

/// Walks an AST and produces a CellValue, resolving cell references via a provider.
public final class FormulaEvaluator {

    public init() {}

    /// Evaluate an AST node to produce a cell value.
    public func evaluate(node: ASTNode, provider: CellValueProvider) -> CellValue {
        switch node {
        case .number(let n):
            return .number(n)

        case .string(let s):
            return .string(s)

        case .boolean(let b):
            return .boolean(b)

        case .error(let e):
            return .error(e)

        case .cellReference(let address):
            return provider.cellValue(at: address)

        case .range:
            // A bare range outside a function is not valid as a scalar value.
            return .error(.value)

        case .unaryOp(let op, let operand):
            return evaluateUnary(op: op, operand: operand, provider: provider)

        case .binaryOp(let op, let left, let right):
            return evaluateBinary(op: op, left: left, right: right, provider: provider)

        case .functionCall(let name, let args):
            return evaluateFunction(name: name, args: args, provider: provider)
        }
    }

    // MARK: - Unary Operations

    private func evaluateUnary(op: UnaryOperator, operand: ASTNode, provider: CellValueProvider) -> CellValue {
        let value = evaluate(node: operand, provider: provider)
        if case .error = value { return value }

        switch op {
        case .negate:
            guard let n = value.numericValue else { return .error(.value) }
            return .number(-n)
        case .percent:
            guard let n = value.numericValue else { return .error(.value) }
            return .number(n / 100.0)
        }
    }

    // MARK: - Binary Operations

    private func evaluateBinary(op: BinaryOperator, left: ASTNode, right: ASTNode, provider: CellValueProvider) -> CellValue {
        let leftVal = evaluate(node: left, provider: provider)
        if case .error = leftVal { return leftVal }
        let rightVal = evaluate(node: right, provider: provider)
        if case .error = rightVal { return rightVal }

        switch op {
        case .add, .subtract, .multiply, .divide, .power:
            return evaluateArithmetic(op: op, left: leftVal, right: rightVal)
        case .concatenate:
            return .string(leftVal.displayString + rightVal.displayString)
        case .equal, .notEqual, .lessThan, .greaterThan, .lessEqual, .greaterEqual:
            return evaluateComparison(op: op, left: leftVal, right: rightVal)
        }
    }

    private func evaluateArithmetic(op: BinaryOperator, left: CellValue, right: CellValue) -> CellValue {
        guard let a = left.numericValue, let b = right.numericValue else {
            // String concatenation with & is not handled here (not in spec),
            // so arithmetic on non-numeric values is an error.
            return .error(.value)
        }

        switch op {
        case .add:      return .number(a + b)
        case .subtract: return .number(a - b)
        case .multiply: return .number(a * b)
        case .divide:
            guard b != 0 else { return .error(.divZero) }
            return .number(a / b)
        case .power:
            let result = pow(a, b)
            guard result.isFinite else { return .error(.num) }
            return .number(result)
        default:
            return .error(.generic)
        }
    }

    private func evaluateComparison(op: BinaryOperator, left: CellValue, right: CellValue) -> CellValue {
        // Compare same-type values; mixed types follow Excel conventions.
        switch (left, right) {
        case (.number(let a), .number(let b)):
            return .boolean(compareOrdered(a, b, op: op))
        case (.string(let a), .string(let b)):
            let cmp = a.localizedCaseInsensitiveCompare(b)
            return .boolean(compareResult(cmp, op: op))
        case (.boolean(let a), .boolean(let b)):
            return .boolean(compareOrdered(a ? 1 : 0, b ? 1 : 0, op: op))
        case (.empty, .empty):
            return .boolean(op == .equal || op == .lessEqual || op == .greaterEqual)
        case (.empty, .number(let n)):
            return .boolean(compareOrdered(0, n, op: op))
        case (.number(let n), .empty):
            return .boolean(compareOrdered(n, 0, op: op))
        case (.empty, .string(let s)):
            return .boolean(compareOrdered("", s, op: op))
        case (.string(let s), .empty):
            return .boolean(compareOrdered(s, "", op: op))
        default:
            // For equality/inequality between different types, they are not equal.
            switch op {
            case .equal:    return .boolean(false)
            case .notEqual: return .boolean(true)
            default:        return .error(.value)
            }
        }
    }

    private func compareOrdered<T: Comparable>(_ a: T, _ b: T, op: BinaryOperator) -> Bool {
        switch op {
        case .equal:        return a == b
        case .notEqual:     return a != b
        case .lessThan:     return a < b
        case .greaterThan:  return a > b
        case .lessEqual:    return a <= b
        case .greaterEqual: return a >= b
        default:            return false
        }
    }

    private func compareResult(_ cmp: ComparisonResult, op: BinaryOperator) -> Bool {
        switch op {
        case .equal:        return cmp == .orderedSame
        case .notEqual:     return cmp != .orderedSame
        case .lessThan:     return cmp == .orderedAscending
        case .greaterThan:  return cmp == .orderedDescending
        case .lessEqual:    return cmp != .orderedDescending
        case .greaterEqual: return cmp != .orderedAscending
        default:            return false
        }
    }

    private func compareOrdered(_ a: String, _ b: String, op: BinaryOperator) -> Bool {
        let cmp = a.localizedCaseInsensitiveCompare(b)
        return compareResult(cmp, op: op)
    }

    // MARK: - Function Evaluation

    private func evaluateFunction(name: String, args: [ASTNode], provider: CellValueProvider) -> CellValue {
        switch name {
        // Aggregate functions
        case "SUM":         return fnSum(args: args, provider: provider)
        case "AVERAGE":     return fnAverage(args: args, provider: provider)
        case "MIN":         return fnMin(args: args, provider: provider)
        case "MAX":         return fnMax(args: args, provider: provider)
        case "COUNT":       return fnCount(args: args, provider: provider)
        case "COUNTA":      return fnCountA(args: args, provider: provider)

        // Logical
        case "IF":          return fnIf(args: args, provider: provider)
        case "AND":         return fnAnd(args: args, provider: provider)
        case "OR":          return fnOr(args: args, provider: provider)
        case "NOT":         return fnNot(args: args, provider: provider)

        // Math
        case "ABS":         return fnAbs(args: args, provider: provider)
        case "ROUND":       return fnRound(args: args, provider: provider)
        case "INT":         return fnInt(args: args, provider: provider)
        case "MOD":         return fnMod(args: args, provider: provider)
        case "POWER":       return fnPowerFn(args: args, provider: provider)
        case "SQRT":        return fnSqrt(args: args, provider: provider)

        // String
        case "LEN":         return fnLen(args: args, provider: provider)
        case "LEFT":        return fnLeft(args: args, provider: provider)
        case "RIGHT":       return fnRight(args: args, provider: provider)
        case "MID":         return fnMid(args: args, provider: provider)
        case "UPPER":       return fnUpper(args: args, provider: provider)
        case "LOWER":       return fnLower(args: args, provider: provider)
        case "TRIM":        return fnTrim(args: args, provider: provider)
        case "CONCATENATE": return fnConcatenate(args: args, provider: provider)
        case "TEXT":        return fnText(args: args, provider: provider)
        case "FIND":        return fnFind(args: args, provider: provider)
        case "SUBSTITUTE":  return fnSubstitute(args: args, provider: provider)

        // Lookup
        case "VLOOKUP":     return fnVlookup(args: args, provider: provider)
        case "INDEX":       return fnIndex(args: args, provider: provider)
        case "MATCH":       return fnMatch(args: args, provider: provider)

        // Conditional
        case "SUMIF":       return fnSumIf(args: args, provider: provider)
        case "COUNTIF":     return fnCountIf(args: args, provider: provider)
        case "IFERROR":     return fnIfError(args: args, provider: provider)
        case "IFNA":        return fnIfNA(args: args, provider: provider)

        // Math (additional)
        case "CEILING":     return fnCeiling(args: args, provider: provider)
        case "FLOOR":       return fnFloor(args: args, provider: provider)
        case "LOG":         return fnLog(args: args, provider: provider)
        case "LN":          return fnLn(args: args, provider: provider)
        case "EXP":         return fnExp(args: args, provider: provider)
        case "TRUNC":       return fnTrunc(args: args, provider: provider)

        // Date/Info
        case "TODAY":       return fnToday(args: args, provider: provider)
        case "NOW":         return fnNow(args: args, provider: provider)
        case "ISBLANK":     return fnIsBlank(args: args, provider: provider)
        case "ISNUMBER":    return fnIsNumber(args: args, provider: provider)
        case "ISTEXT":      return fnIsText(args: args, provider: provider)
        case "ISERROR":     return fnIsError(args: args, provider: provider)
        case "TYPE":        return fnType(args: args, provider: provider)

        default:
            return .error(.name)
        }
    }

    // MARK: - Expand Ranges

    /// Expand an argument list, flattening ranges into individual cell values.
    private func expandArguments(_ args: [ASTNode], provider: CellValueProvider) -> [CellValue] {
        var result: [CellValue] = []
        for arg in args {
            switch arg {
            case .range(let start, let end):
                let range = CellRange(start: start, end: end)
                for addr in range.allAddresses {
                    result.append(provider.cellValue(at: addr))
                }
            default:
                result.append(evaluate(node: arg, provider: provider))
            }
        }
        return result
    }

    /// Result type for numeric extraction that avoids requiring Error conformance.
    private enum NumericResult {
        case success([Double])
        case failure(CellValue)
    }

    /// Extract numeric values from expanded arguments, propagating errors.
    private func numericValues(from args: [ASTNode], provider: CellValueProvider) -> NumericResult {
        let expanded = expandArguments(args, provider: provider)
        var numbers: [Double] = []
        for value in expanded {
            switch value {
            case .error:
                return .failure(value)
            case .number(let n):
                numbers.append(n)
            case .boolean(let b):
                numbers.append(b ? 1 : 0)
            case .empty, .string, .date:
                // Aggregate functions skip non-numeric values
                continue
            }
        }
        return .success(numbers)
    }

    // MARK: - Aggregate Functions

    private func fnSum(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        switch numericValues(from: args, provider: provider) {
        case .failure(let err): return err
        case .success(let nums): return .number(nums.reduce(0, +))
        }
    }

    private func fnAverage(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        switch numericValues(from: args, provider: provider) {
        case .failure(let err): return err
        case .success(let nums):
            guard !nums.isEmpty else { return .error(.divZero) }
            return .number(nums.reduce(0, +) / Double(nums.count))
        }
    }

    private func fnMin(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        switch numericValues(from: args, provider: provider) {
        case .failure(let err): return err
        case .success(let nums):
            guard let m = nums.min() else { return .number(0) }
            return .number(m)
        }
    }

    private func fnMax(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        switch numericValues(from: args, provider: provider) {
        case .failure(let err): return err
        case .success(let nums):
            guard let m = nums.max() else { return .number(0) }
            return .number(m)
        }
    }

    private func fnCount(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        let expanded = expandArguments(args, provider: provider)
        var count = 0
        for value in expanded {
            if case .error = value { return value }
            if value.numericValue != nil { count += 1 }
        }
        return .number(Double(count))
    }

    private func fnCountA(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        let expanded = expandArguments(args, provider: provider)
        var count = 0
        for value in expanded {
            if case .error = value { return value }
            if !value.isEmpty { count += 1 }
        }
        return .number(Double(count))
    }

    // MARK: - Logical Functions

    private func fnIf(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count >= 2 && args.count <= 3 else { return .error(.value) }

        let condition = evaluate(node: args[0], provider: provider)
        if case .error = condition { return condition }

        let isTruthy: Bool
        switch condition {
        case .boolean(let b): isTruthy = b
        case .number(let n):  isTruthy = n != 0
        case .string(let s):  isTruthy = !s.isEmpty
        case .empty:          isTruthy = false
        default:              return .error(.value)
        }

        if isTruthy {
            return evaluate(node: args[1], provider: provider)
        } else if args.count == 3 {
            return evaluate(node: args[2], provider: provider)
        } else {
            return .boolean(false)
        }
    }

    private func fnAnd(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard !args.isEmpty else { return .error(.value) }
        let expanded = expandArguments(args, provider: provider)
        var foundValue = false
        for value in expanded {
            switch value {
            case .error: return value
            case .boolean(let b):
                if !b { return .boolean(false) }
                foundValue = true
            case .number(let n):
                if n == 0 { return .boolean(false) }
                foundValue = true
            case .empty:
                continue
            default:
                return .error(.value)
            }
        }
        guard foundValue else { return .error(.value) }
        return .boolean(true)
    }

    private func fnOr(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard !args.isEmpty else { return .error(.value) }
        let expanded = expandArguments(args, provider: provider)
        var foundValue = false
        for value in expanded {
            switch value {
            case .error: return value
            case .boolean(let b):
                if b { return .boolean(true) }
                foundValue = true
            case .number(let n):
                if n != 0 { return .boolean(true) }
                foundValue = true
            case .empty:
                continue
            default:
                return .error(.value)
            }
        }
        guard foundValue else { return .error(.value) }
        return .boolean(false)
    }

    private func fnNot(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 1 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        switch value {
        case .error: return value
        case .boolean(let b): return .boolean(!b)
        case .number(let n):  return .boolean(n == 0)
        default: return .error(.value)
        }
    }

    // MARK: - Math Functions

    private func fnAbs(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 1 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        if case .error = value { return value }
        guard let n = value.numericValue else { return .error(.value) }
        return .number(abs(n))
    }

    private func fnRound(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 2 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        if case .error = value { return value }
        let digitsVal = evaluate(node: args[1], provider: provider)
        if case .error = digitsVal { return digitsVal }
        guard let n = value.numericValue, let d = digitsVal.numericValue else {
            return .error(.value)
        }
        let digits = Int(d)
        let multiplier = pow(10.0, Double(digits))
        return .number((n * multiplier).rounded() / multiplier)
    }

    private func fnInt(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 1 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        if case .error = value { return value }
        guard let n = value.numericValue else { return .error(.value) }
        return .number(floor(n))
    }

    private func fnMod(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 2 else { return .error(.value) }
        let numVal = evaluate(node: args[0], provider: provider)
        if case .error = numVal { return numVal }
        let divVal = evaluate(node: args[1], provider: provider)
        if case .error = divVal { return divVal }
        guard let n = numVal.numericValue, let d = divVal.numericValue else {
            return .error(.value)
        }
        guard d != 0 else { return .error(.divZero) }
        // Excel MOD: result has the sign of the divisor
        let result = n - d * floor(n / d)
        return .number(result)
    }

    private func fnPowerFn(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 2 else { return .error(.value) }
        let baseVal = evaluate(node: args[0], provider: provider)
        if case .error = baseVal { return baseVal }
        let expVal = evaluate(node: args[1], provider: provider)
        if case .error = expVal { return expVal }
        guard let b = baseVal.numericValue, let e = expVal.numericValue else {
            return .error(.value)
        }
        let result = pow(b, e)
        guard result.isFinite else { return .error(.num) }
        return .number(result)
    }

    private func fnSqrt(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 1 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        if case .error = value { return value }
        guard let n = value.numericValue else { return .error(.value) }
        guard n >= 0 else { return .error(.num) }
        return .number(sqrt(n))
    }

    // MARK: - String Functions

    private func fnLen(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 1 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        if case .error = value { return value }
        let text = coerceToString(value)
        return .number(Double(text.count))
    }

    private func fnLeft(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count >= 1 && args.count <= 2 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        if case .error = value { return value }
        let text = coerceToString(value)

        var numChars = 1
        if args.count == 2 {
            let nVal = evaluate(node: args[1], provider: provider)
            if case .error = nVal { return nVal }
            guard let n = nVal.numericValue else { return .error(.value) }
            guard n >= 0 else { return .error(.value) }
            numChars = Int(n)
        }

        let end = text.index(text.startIndex, offsetBy: min(numChars, text.count))
        return .string(String(text[text.startIndex..<end]))
    }

    private func fnRight(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count >= 1 && args.count <= 2 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        if case .error = value { return value }
        let text = coerceToString(value)

        var numChars = 1
        if args.count == 2 {
            let nVal = evaluate(node: args[1], provider: provider)
            if case .error = nVal { return nVal }
            guard let n = nVal.numericValue else { return .error(.value) }
            guard n >= 0 else { return .error(.value) }
            numChars = Int(n)
        }

        let start = text.index(text.endIndex, offsetBy: -min(numChars, text.count))
        return .string(String(text[start..<text.endIndex]))
    }

    private func fnMid(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 3 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        if case .error = value { return value }
        let text = coerceToString(value)

        let startVal = evaluate(node: args[1], provider: provider)
        if case .error = startVal { return startVal }
        let numVal = evaluate(node: args[2], provider: provider)
        if case .error = numVal { return numVal }

        guard let startNum = startVal.numericValue, let numChars = numVal.numericValue else {
            return .error(.value)
        }
        guard startNum >= 1 && numChars >= 0 else { return .error(.value) }

        let startIndex = Int(startNum) - 1  // 1-based to 0-based
        guard startIndex < text.count else { return .string("") }

        let from = text.index(text.startIndex, offsetBy: startIndex)
        let to = text.index(from, offsetBy: min(Int(numChars), text.count - startIndex))
        return .string(String(text[from..<to]))
    }

    private func fnUpper(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 1 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        if case .error = value { return value }
        return .string(coerceToString(value).uppercased())
    }

    private func fnLower(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 1 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        if case .error = value { return value }
        return .string(coerceToString(value).lowercased())
    }

    private func fnTrim(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 1 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        if case .error = value { return value }
        let text = coerceToString(value)
        // Excel TRIM: removes leading/trailing spaces and collapses internal runs of spaces to one.
        let components = text.split(separator: " ", omittingEmptySubsequences: true)
        return .string(components.joined(separator: " "))
    }

    private func fnConcatenate(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        var result = ""
        for arg in args {
            let value = evaluate(node: arg, provider: provider)
            if case .error = value { return value }
            result += coerceToString(value)
        }
        return .string(result)
    }

    // MARK: - Coercion

    /// Convert a CellValue to its string representation for string functions.
    private func coerceToString(_ value: CellValue) -> String {
        switch value {
        case .string(let s): return s
        case .number: return value.displayString
        case .boolean(let b): return b ? "TRUE" : "FALSE"
        case .empty: return ""
        case .date: return value.displayString
        case .error(let e): return e.rawValue
        }
    }

    // MARK: - Lookup Functions

    /// Expand a range argument into a 2D grid of cell values (rows x columns).
    private func expandRangeToGrid(_ node: ASTNode, provider: CellValueProvider) -> (values: [[CellValue]], rowCount: Int, colCount: Int)? {
        guard case .range(let start, let end) = node else { return nil }
        let range = CellRange(start: start, end: end)
        var grid: [[CellValue]] = []
        for r in range.start.row...range.end.row {
            var row: [CellValue] = []
            for c in range.start.column...range.end.column {
                row.append(provider.cellValue(at: CellAddress(column: c, row: r)))
            }
            grid.append(row)
        }
        return (grid, range.rowCount, range.columnCount)
    }

    /// Expand a range into a flat list of cell values (row-major order).
    private func expandRangeToList(_ node: ASTNode, provider: CellValueProvider) -> [CellValue]? {
        guard case .range(let start, let end) = node else { return nil }
        let range = CellRange(start: start, end: end)
        return range.allAddresses.map { provider.cellValue(at: $0) }
    }

    /// Check whether a CellValue matches a criteria string.
    /// Supports: exact match, comparison operators (">5", "<=10", "<>text"), and wildcard * and ?.
    private func matchesCriteria(_ value: CellValue, criteria: String) -> Bool {
        // Check for comparison operator prefixes
        if criteria.hasPrefix(">=") {
            let operand = String(criteria.dropFirst(2))
            return compareCriteria(value, op: .greaterEqual, operand: operand)
        } else if criteria.hasPrefix("<=") {
            let operand = String(criteria.dropFirst(2))
            return compareCriteria(value, op: .lessEqual, operand: operand)
        } else if criteria.hasPrefix("<>") {
            let operand = String(criteria.dropFirst(2))
            return compareCriteria(value, op: .notEqual, operand: operand)
        } else if criteria.hasPrefix(">") {
            let operand = String(criteria.dropFirst(1))
            return compareCriteria(value, op: .greaterThan, operand: operand)
        } else if criteria.hasPrefix("<") {
            let operand = String(criteria.dropFirst(1))
            return compareCriteria(value, op: .lessThan, operand: operand)
        } else if criteria.hasPrefix("=") {
            let operand = String(criteria.dropFirst(1))
            return compareCriteria(value, op: .equal, operand: operand)
        }

        // Exact match (case-insensitive for strings)
        if let criteriaNum = Double(criteria) {
            if let valNum = value.numericValue {
                return valNum == criteriaNum
            }
            return false
        }
        // String comparison (case-insensitive)
        let valueStr = coerceToString(value)
        return valueStr.caseInsensitiveCompare(criteria) == .orderedSame
    }

    private func compareCriteria(_ value: CellValue, op: BinaryOperator, operand: String) -> Bool {
        if let operandNum = Double(operand), let valNum = value.numericValue {
            return compareOrdered(valNum, operandNum, op: op)
        }
        let valueStr = coerceToString(value)
        let cmp = valueStr.localizedCaseInsensitiveCompare(operand)
        return compareResult(cmp, op: op)
    }

    private func fnVlookup(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count >= 3 && args.count <= 4 else { return .error(.value) }

        let lookupVal = evaluate(node: args[0], provider: provider)
        if case .error = lookupVal { return lookupVal }

        guard let grid = expandRangeToGrid(args[1], provider: provider) else {
            return .error(.value)
        }

        let colIndexVal = evaluate(node: args[2], provider: provider)
        if case .error = colIndexVal { return colIndexVal }
        guard let colIndexDbl = colIndexVal.numericValue else { return .error(.value) }
        let colIndex = Int(colIndexDbl)
        guard colIndex >= 1 && colIndex <= grid.colCount else { return .error(.ref) }

        var rangeLookup = true
        if args.count == 4 {
            let rlVal = evaluate(node: args[3], provider: provider)
            if case .error = rlVal { return rlVal }
            switch rlVal {
            case .boolean(let b): rangeLookup = b
            case .number(let n): rangeLookup = n != 0
            default: rangeLookup = true
            }
        }

        if !rangeLookup {
            // Exact match: search first column for lookupVal
            for r in 0..<grid.rowCount {
                if cellValuesEqual(grid.values[r][0], lookupVal) {
                    return grid.values[r][colIndex - 1]
                }
            }
            return .error(.na)
        } else {
            // Approximate match: first column must be sorted ascending.
            // Find largest value <= lookupVal.
            guard let lookupNum = lookupVal.numericValue else {
                // For strings, do case-insensitive comparison
                var bestRow: Int? = nil
                let lookupStr = coerceToString(lookupVal)
                for r in 0..<grid.rowCount {
                    let cellStr = coerceToString(grid.values[r][0])
                    if cellStr.localizedCaseInsensitiveCompare(lookupStr) != .orderedDescending {
                        bestRow = r
                    } else {
                        break
                    }
                }
                guard let row = bestRow else { return .error(.na) }
                return grid.values[row][colIndex - 1]
            }
            var bestRow: Int? = nil
            for r in 0..<grid.rowCount {
                guard let cellNum = grid.values[r][0].numericValue else { continue }
                if cellNum <= lookupNum {
                    bestRow = r
                } else {
                    break
                }
            }
            guard let row = bestRow else { return .error(.na) }
            return grid.values[row][colIndex - 1]
        }
    }

    /// Helper: check if two CellValues are equal (case-insensitive for strings).
    private func cellValuesEqual(_ a: CellValue, _ b: CellValue) -> Bool {
        switch (a, b) {
        case (.number(let x), .number(let y)): return x == y
        case (.string(let x), .string(let y)):
            return x.caseInsensitiveCompare(y) == .orderedSame
        case (.boolean(let x), .boolean(let y)): return x == y
        case (.empty, .empty): return true
        case (.number(let n), .empty): return n == 0
        case (.empty, .number(let n)): return n == 0
        case (.number(let n), .boolean(let b)): return n == (b ? 1 : 0)
        case (.boolean(let b), .number(let n)): return n == (b ? 1 : 0)
        default: return false
        }
    }

    private func fnIndex(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count >= 2 && args.count <= 3 else { return .error(.value) }

        guard let grid = expandRangeToGrid(args[0], provider: provider) else {
            return .error(.value)
        }

        let rowNumVal = evaluate(node: args[1], provider: provider)
        if case .error = rowNumVal { return rowNumVal }
        guard let rowNumDbl = rowNumVal.numericValue else { return .error(.value) }
        let rowNum = Int(rowNumDbl)

        var colNum = 1
        if args.count == 3 {
            let colNumVal = evaluate(node: args[2], provider: provider)
            if case .error = colNumVal { return colNumVal }
            guard let colNumDbl = colNumVal.numericValue else { return .error(.value) }
            colNum = Int(colNumDbl)
        }

        guard rowNum >= 1 && rowNum <= grid.rowCount else { return .error(.ref) }
        guard colNum >= 1 && colNum <= grid.colCount else { return .error(.ref) }

        return grid.values[rowNum - 1][colNum - 1]
    }

    private func fnMatch(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count >= 2 && args.count <= 3 else { return .error(.value) }

        let lookupVal = evaluate(node: args[0], provider: provider)
        if case .error = lookupVal { return lookupVal }

        guard let list = expandRangeToList(args[1], provider: provider) else {
            return .error(.value)
        }

        var matchType = 1
        if args.count == 3 {
            let mtVal = evaluate(node: args[2], provider: provider)
            if case .error = mtVal { return mtVal }
            guard let mtNum = mtVal.numericValue else { return .error(.value) }
            matchType = Int(mtNum)
        }

        if matchType == 0 {
            // Exact match
            for (i, val) in list.enumerated() {
                if cellValuesEqual(val, lookupVal) {
                    return .number(Double(i + 1))
                }
            }
            return .error(.na)
        } else if matchType == 1 {
            // Largest value <= lookupVal (data must be ascending)
            guard let lookupNum = lookupVal.numericValue else {
                // String match
                let lookupStr = coerceToString(lookupVal)
                var bestIndex: Int? = nil
                for (i, val) in list.enumerated() {
                    let valStr = coerceToString(val)
                    if valStr.localizedCaseInsensitiveCompare(lookupStr) != .orderedDescending {
                        bestIndex = i
                    } else {
                        break
                    }
                }
                guard let idx = bestIndex else { return .error(.na) }
                return .number(Double(idx + 1))
            }
            var bestIndex: Int? = nil
            for (i, val) in list.enumerated() {
                guard let valNum = val.numericValue else { continue }
                if valNum <= lookupNum {
                    bestIndex = i
                } else {
                    break
                }
            }
            guard let idx = bestIndex else { return .error(.na) }
            return .number(Double(idx + 1))
        } else {
            // matchType == -1: Smallest value >= lookupVal (data must be descending)
            guard let lookupNum = lookupVal.numericValue else {
                let lookupStr = coerceToString(lookupVal)
                var bestIndex: Int? = nil
                for (i, val) in list.enumerated() {
                    let valStr = coerceToString(val)
                    if valStr.localizedCaseInsensitiveCompare(lookupStr) != .orderedAscending {
                        bestIndex = i
                    } else {
                        break
                    }
                }
                guard let idx = bestIndex else { return .error(.na) }
                return .number(Double(idx + 1))
            }
            var bestIndex: Int? = nil
            for (i, val) in list.enumerated() {
                guard let valNum = val.numericValue else { continue }
                if valNum >= lookupNum {
                    bestIndex = i
                } else {
                    break
                }
            }
            guard let idx = bestIndex else { return .error(.na) }
            return .number(Double(idx + 1))
        }
    }

    // MARK: - Conditional Functions

    private func fnSumIf(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count >= 2 && args.count <= 3 else { return .error(.value) }

        guard case .range(let rangeStart, let rangeEnd) = args[0] else {
            return .error(.value)
        }

        let criteriaVal = evaluate(node: args[1], provider: provider)
        if case .error = criteriaVal { return criteriaVal }
        let criteria = coerceToString(criteriaVal)

        let criteriaRange = CellRange(start: rangeStart, end: rangeEnd)
        let criteriaAddresses = criteriaRange.allAddresses

        // Determine sum range
        let sumAddresses: [CellAddress]
        if args.count == 3 {
            guard case .range(let sumStart, let sumEnd) = args[2] else {
                return .error(.value)
            }
            let sumRange = CellRange(start: sumStart, end: sumEnd)
            sumAddresses = sumRange.allAddresses
        } else {
            sumAddresses = criteriaAddresses
        }

        var total = 0.0
        for (i, addr) in criteriaAddresses.enumerated() {
            let cellVal = provider.cellValue(at: addr)
            if matchesCriteria(cellVal, criteria: criteria) {
                if i < sumAddresses.count {
                    let sumVal = provider.cellValue(at: sumAddresses[i])
                    if let n = sumVal.numericValue {
                        total += n
                    }
                }
            }
        }
        return .number(total)
    }

    private func fnCountIf(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 2 else { return .error(.value) }

        guard case .range(let rangeStart, let rangeEnd) = args[0] else {
            return .error(.value)
        }

        let criteriaVal = evaluate(node: args[1], provider: provider)
        if case .error = criteriaVal { return criteriaVal }
        let criteria = coerceToString(criteriaVal)

        let range = CellRange(start: rangeStart, end: rangeEnd)
        var count = 0
        for addr in range.allAddresses {
            let cellVal = provider.cellValue(at: addr)
            if matchesCriteria(cellVal, criteria: criteria) {
                count += 1
            }
        }
        return .number(Double(count))
    }

    private func fnIfError(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 2 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        if case .error = value {
            return evaluate(node: args[1], provider: provider)
        }
        return value
    }

    private func fnIfNA(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 2 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        if case .error(.na) = value {
            return evaluate(node: args[1], provider: provider)
        }
        return value
    }

    // MARK: - Additional Math Functions

    private func fnCeiling(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 2 else { return .error(.value) }
        let numVal = evaluate(node: args[0], provider: provider)
        if case .error = numVal { return numVal }
        let sigVal = evaluate(node: args[1], provider: provider)
        if case .error = sigVal { return sigVal }
        guard let number = numVal.numericValue, let significance = sigVal.numericValue else {
            return .error(.value)
        }
        guard significance != 0 else { return .number(0) }
        // Signs must match in Excel CEILING
        guard (number >= 0 && significance > 0) || (number <= 0 && significance < 0) else {
            return .error(.num)
        }
        return .number(ceil(number / significance) * significance)
    }

    private func fnFloor(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 2 else { return .error(.value) }
        let numVal = evaluate(node: args[0], provider: provider)
        if case .error = numVal { return numVal }
        let sigVal = evaluate(node: args[1], provider: provider)
        if case .error = sigVal { return sigVal }
        guard let number = numVal.numericValue, let significance = sigVal.numericValue else {
            return .error(.value)
        }
        guard significance != 0 else { return .error(.divZero) }
        guard (number >= 0 && significance > 0) || (number <= 0 && significance < 0) else {
            return .error(.num)
        }
        return .number(Foundation.floor(number / significance) * significance)
    }

    private func fnLog(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count >= 1 && args.count <= 2 else { return .error(.value) }
        let numVal = evaluate(node: args[0], provider: provider)
        if case .error = numVal { return numVal }
        guard let number = numVal.numericValue else { return .error(.value) }
        guard number > 0 else { return .error(.num) }

        var base = 10.0
        if args.count == 2 {
            let baseVal = evaluate(node: args[1], provider: provider)
            if case .error = baseVal { return baseVal }
            guard let b = baseVal.numericValue else { return .error(.value) }
            guard b > 0 && b != 1 else { return .error(.num) }
            base = b
        }
        let result = Foundation.log(number) / Foundation.log(base)
        guard result.isFinite else { return .error(.num) }
        return .number(result)
    }

    private func fnLn(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 1 else { return .error(.value) }
        let numVal = evaluate(node: args[0], provider: provider)
        if case .error = numVal { return numVal }
        guard let number = numVal.numericValue else { return .error(.value) }
        guard number > 0 else { return .error(.num) }
        return .number(Foundation.log(number))
    }

    private func fnExp(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 1 else { return .error(.value) }
        let numVal = evaluate(node: args[0], provider: provider)
        if case .error = numVal { return numVal }
        guard let number = numVal.numericValue else { return .error(.value) }
        let result = Foundation.exp(number)
        guard result.isFinite else { return .error(.num) }
        return .number(result)
    }

    private func fnTrunc(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count >= 1 && args.count <= 2 else { return .error(.value) }
        let numVal = evaluate(node: args[0], provider: provider)
        if case .error = numVal { return numVal }
        guard let number = numVal.numericValue else { return .error(.value) }

        var numDigits = 0
        if args.count == 2 {
            let digitsVal = evaluate(node: args[1], provider: provider)
            if case .error = digitsVal { return digitsVal }
            guard let d = digitsVal.numericValue else { return .error(.value) }
            numDigits = Int(d)
        }

        let multiplier = pow(10.0, Double(numDigits))
        let result = (number * multiplier).rounded(.towardZero) / multiplier
        return .number(result)
    }

    // MARK: - Additional Text Functions

    private func fnText(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 2 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        if case .error = value { return value }
        let fmtVal = evaluate(node: args[1], provider: provider)
        if case .error = fmtVal { return fmtVal }
        guard case .string(let formatText) = fmtVal else { return .error(.value) }

        guard let number = value.numericValue else {
            return .string(coerceToString(value))
        }

        // Basic format support: "0.00", "#,##0", "0%", "#,##0.00"
        if formatText == "0%" {
            return .string(String(format: "%.0f%%", number * 100))
        } else if formatText == "0.00%" {
            return .string(String(format: "%.2f%%", number * 100))
        }

        // Count decimal places from format
        let isPercentage = formatText.hasSuffix("%")
        let fmt = isPercentage ? String(formatText.dropLast()) : formatText
        let hasThousandsSep = fmt.contains(",")

        // Count digits after decimal point
        var decimals = 0
        if let dotIndex = fmt.lastIndex(of: ".") {
            let afterDot = fmt[fmt.index(after: dotIndex)...]
            decimals = afterDot.filter { $0 == "0" || $0 == "#" }.count
        }

        let displayNum = isPercentage ? number * 100 : number
        var formatted: String

        if hasThousandsSep {
            let nf = NumberFormatter()
            nf.numberStyle = .decimal
            nf.minimumFractionDigits = decimals
            nf.maximumFractionDigits = decimals
            nf.groupingSeparator = ","
            nf.decimalSeparator = "."
            nf.usesGroupingSeparator = true
            formatted = nf.string(from: NSNumber(value: displayNum)) ?? String(format: "%.\(decimals)f", displayNum)
        } else {
            formatted = String(format: "%.\(decimals)f", displayNum)
        }

        if isPercentage {
            formatted += "%"
        }
        return .string(formatted)
    }

    private func fnFind(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count >= 2 && args.count <= 3 else { return .error(.value) }

        let findVal = evaluate(node: args[0], provider: provider)
        if case .error = findVal { return findVal }
        let findText = coerceToString(findVal)

        let withinVal = evaluate(node: args[1], provider: provider)
        if case .error = withinVal { return withinVal }
        let withinText = coerceToString(withinVal)

        var startNum = 1
        if args.count == 3 {
            let startVal = evaluate(node: args[2], provider: provider)
            if case .error = startVal { return startVal }
            guard let s = startVal.numericValue else { return .error(.value) }
            guard s >= 1 else { return .error(.value) }
            startNum = Int(s)
        }

        guard startNum <= withinText.count + 1 else { return .error(.value) }

        let searchStart = withinText.index(withinText.startIndex, offsetBy: startNum - 1)
        let searchRange = searchStart..<withinText.endIndex

        // Case-sensitive search
        guard let foundRange = withinText.range(of: findText, range: searchRange) else {
            return .error(.value)
        }

        let position = withinText.distance(from: withinText.startIndex, to: foundRange.lowerBound) + 1
        return .number(Double(position))
    }

    private func fnSubstitute(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count >= 3 && args.count <= 4 else { return .error(.value) }

        let textVal = evaluate(node: args[0], provider: provider)
        if case .error = textVal { return textVal }
        let text = coerceToString(textVal)

        let oldVal = evaluate(node: args[1], provider: provider)
        if case .error = oldVal { return oldVal }
        let oldText = coerceToString(oldVal)

        let newVal = evaluate(node: args[2], provider: provider)
        if case .error = newVal { return newVal }
        let newText = coerceToString(newVal)

        guard !oldText.isEmpty else { return .string(text) }

        if args.count == 4 {
            let instanceVal = evaluate(node: args[3], provider: provider)
            if case .error = instanceVal { return instanceVal }
            guard let instNum = instanceVal.numericValue else { return .error(.value) }
            let instanceNum = Int(instNum)
            guard instanceNum >= 1 else { return .error(.value) }

            // Replace only the nth occurrence
            var result = text
            var occurrenceCount = 0
            var searchStart = result.startIndex
            while let range = result.range(of: oldText, range: searchStart..<result.endIndex) {
                occurrenceCount += 1
                if occurrenceCount == instanceNum {
                    result.replaceSubrange(range, with: newText)
                    break
                }
                searchStart = range.upperBound
            }
            return .string(result)
        } else {
            // Replace all occurrences
            return .string(text.replacingOccurrences(of: oldText, with: newText))
        }
    }

    // MARK: - Date/Info Functions

    private func fnToday(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.isEmpty else { return .error(.value) }
        // Return as a date serial number (days since 1899-12-30 per Excel convention)
        let now = Date()
        let calendar = Calendar.current
        let components = calendar.dateComponents([.year, .month, .day], from: now)
        guard let dateOnly = calendar.date(from: components) else { return .error(.value) }
        let referenceDate = excelEpoch()
        let days = calendar.dateComponents([.day], from: referenceDate, to: dateOnly).day ?? 0
        return .number(Double(days))
    }

    private func fnNow(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.isEmpty else { return .error(.value) }
        let now = Date()
        let referenceDate = excelEpoch()
        let interval = now.timeIntervalSince(referenceDate)
        let days = interval / 86400.0
        return .number(days)
    }

    /// Excel epoch: 1899-12-30 (day 0).
    private func excelEpoch() -> Date {
        var comps = DateComponents()
        comps.year = 1899
        comps.month = 12
        comps.day = 30
        comps.hour = 0
        comps.minute = 0
        comps.second = 0
        let cal = Calendar(identifier: .gregorian)
        return cal.date(from: comps)!
    }

    private func fnIsBlank(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 1 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        return .boolean(value.isEmpty)
    }

    private func fnIsNumber(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 1 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        if case .number = value { return .boolean(true) }
        return .boolean(false)
    }

    private func fnIsText(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 1 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        if case .string = value { return .boolean(true) }
        return .boolean(false)
    }

    private func fnIsError(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 1 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        return .boolean(value.isError)
    }

    private func fnType(args: [ASTNode], provider: CellValueProvider) -> CellValue {
        guard args.count == 1 else { return .error(.value) }
        let value = evaluate(node: args[0], provider: provider)
        switch value {
        case .number:  return .number(1)
        case .string:  return .number(2)
        case .boolean: return .number(4)
        case .error:   return .number(16)
        case .empty:   return .number(1)   // Excel treats blank as number type
        case .date:    return .number(1)   // Dates are numbers in Excel
        }
    }
}
