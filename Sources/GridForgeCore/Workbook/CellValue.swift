import Foundation

/// Represents the semantic value of a spreadsheet cell
public enum CellValue: Equatable, Sendable {
    case empty
    case string(String)
    case number(Double)
    case boolean(Bool)
    case date(Date)
    case error(CellError)

    public var displayString: String {
        switch self {
        case .empty: return ""
        case .string(let s): return s
        case .number(let n):
            if n == n.rounded() && abs(n) < 1e15 {
                return String(format: "%.0f", n)
            }
            return String(n)
        case .boolean(let b): return b ? "TRUE" : "FALSE"
        case .date(let d):
            let f = DateFormatter()
            f.dateStyle = .short
            return f.string(from: d)
        case .error(let e): return e.rawValue
        }
    }

    public var numericValue: Double? {
        switch self {
        case .number(let n): return n
        case .boolean(let b): return b ? 1 : 0
        case .empty: return 0
        default: return nil
        }
    }

    public var isError: Bool {
        if case .error = self { return true }
        return false
    }

    public var isEmpty: Bool {
        if case .empty = self { return true }
        return false
    }
}

/// Spreadsheet error types matching Excel conventions
public enum CellError: String, Equatable, Sendable, CaseIterable {
    case value = "#VALUE!"
    case ref = "#REF!"
    case divZero = "#DIV/0!"
    case name = "#NAME?"
    case na = "#N/A"
    case circular = "#CIRCULAR!"
    case generic = "#ERROR!"
    case syntax = "#SYNTAX!"
    case num = "#NUM!"
}
