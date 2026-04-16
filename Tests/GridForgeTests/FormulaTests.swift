import XCTest
@testable import GridForgeCore

final class FormulaTests: XCTestCase {

    var engine: FormulaEngine!
    var sheet: Worksheet!

    override func setUp() {
        super.setUp()
        engine = FormulaEngine()
        sheet = Worksheet(name: "Test")
    }

    // MARK: - Basic Values

    func testNumberLiteral() {
        let result = engine.evaluate(formula: "42", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(result, .number(42))
    }

    func testDecimalLiteral() {
        let result = engine.evaluate(formula: "3.14", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(result, .number(3.14))
    }

    func testStringLiteral() {
        let result = engine.evaluate(formula: "\"hello\"", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(result, .string("hello"))
    }

    func testBooleanLiteral() {
        let t = engine.evaluate(formula: "TRUE", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(t, .boolean(true))

        let f = engine.evaluate(formula: "FALSE", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(f, .boolean(false))
    }

    // MARK: - Arithmetic

    func testAddition() {
        let result = engine.evaluate(formula: "1+2", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(result, .number(3))
    }

    func testSubtraction() {
        let result = engine.evaluate(formula: "10-3", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(result, .number(7))
    }

    func testMultiplication() {
        let result = engine.evaluate(formula: "4*5", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(result, .number(20))
    }

    func testDivision() {
        let result = engine.evaluate(formula: "10/4", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(result, .number(2.5))
    }

    func testDivisionByZero() {
        let result = engine.evaluate(formula: "1/0", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(result, .error(.divZero))
    }

    func testPrecedence() {
        let result = engine.evaluate(formula: "2+3*4", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(result, .number(14))
    }

    func testParentheses() {
        let result = engine.evaluate(formula: "(2+3)*4", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(result, .number(20))
    }

    func testNegation() {
        let result = engine.evaluate(formula: "-5", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(result, .number(-5))
    }

    func testComplexExpression() {
        let result = engine.evaluate(formula: "(10+5)*2-3", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(result, .number(27))
    }

    // MARK: - Cell References

    func testCellReference() {
        sheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "42")
        let result = engine.evaluate(formula: "A1", in: sheet, at: CellAddress(column: 1, row: 0))
        XCTAssertEqual(result, .number(42))
    }

    func testCellReferenceInExpression() {
        sheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "10")
        sheet.setRawInput(at: CellAddress(column: 1, row: 0), rawInput: "20")
        let result = engine.evaluate(formula: "A1+B1", in: sheet, at: CellAddress(column: 2, row: 0))
        XCTAssertEqual(result, .number(30))
    }

    func testEmptyCellReference() {
        // Empty cell referenced alone returns empty (or 0 in numeric context)
        let result = engine.evaluate(formula: "A1", in: sheet, at: CellAddress(column: 1, row: 0))
        // Evaluator returns .empty for empty cells; in arithmetic context they become 0
        XCTAssertTrue(result == .empty || result == .number(0))

        // But in arithmetic, empty is treated as 0
        let sum = engine.evaluate(formula: "A1+1", in: sheet, at: CellAddress(column: 1, row: 0))
        XCTAssertEqual(sum, .number(1))
    }

    // MARK: - Functions

    func testSUM() {
        sheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "1")
        sheet.setRawInput(at: CellAddress(column: 0, row: 1), rawInput: "2")
        sheet.setRawInput(at: CellAddress(column: 0, row: 2), rawInput: "3")

        let result = engine.evaluate(formula: "SUM(A1:A3)", in: sheet, at: CellAddress(column: 1, row: 0))
        XCTAssertEqual(result, .number(6))
    }

    func testAVERAGE() {
        sheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "10")
        sheet.setRawInput(at: CellAddress(column: 0, row: 1), rawInput: "20")
        sheet.setRawInput(at: CellAddress(column: 0, row: 2), rawInput: "30")

        let result = engine.evaluate(formula: "AVERAGE(A1:A3)", in: sheet, at: CellAddress(column: 1, row: 0))
        XCTAssertEqual(result, .number(20))
    }

    func testMIN() {
        sheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "5")
        sheet.setRawInput(at: CellAddress(column: 0, row: 1), rawInput: "3")
        sheet.setRawInput(at: CellAddress(column: 0, row: 2), rawInput: "8")

        let result = engine.evaluate(formula: "MIN(A1:A3)", in: sheet, at: CellAddress(column: 1, row: 0))
        XCTAssertEqual(result, .number(3))
    }

    func testMAX() {
        sheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "5")
        sheet.setRawInput(at: CellAddress(column: 0, row: 1), rawInput: "3")
        sheet.setRawInput(at: CellAddress(column: 0, row: 2), rawInput: "8")

        let result = engine.evaluate(formula: "MAX(A1:A3)", in: sheet, at: CellAddress(column: 1, row: 0))
        XCTAssertEqual(result, .number(8))
    }

    func testCOUNT() {
        sheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "1")
        sheet.setRawInput(at: CellAddress(column: 0, row: 1), rawInput: "hello")
        sheet.setRawInput(at: CellAddress(column: 0, row: 2), rawInput: "3")

        let result = engine.evaluate(formula: "COUNT(A1:A3)", in: sheet, at: CellAddress(column: 1, row: 0))
        XCTAssertEqual(result, .number(2))
    }

    func testIF() {
        let t = engine.evaluate(formula: "IF(TRUE,\"yes\",\"no\")", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(t, .string("yes"))

        let f = engine.evaluate(formula: "IF(FALSE,\"yes\",\"no\")", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(f, .string("no"))
    }

    func testIFWithComparison() {
        sheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "10")
        let result = engine.evaluate(formula: "IF(A1>5,\"big\",\"small\")", in: sheet, at: CellAddress(column: 1, row: 0))
        XCTAssertEqual(result, .string("big"))
    }

    func testSUMMultipleArgs() {
        sheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "1")
        sheet.setRawInput(at: CellAddress(column: 1, row: 0), rawInput: "2")
        let result = engine.evaluate(formula: "SUM(A1,B1,10)", in: sheet, at: CellAddress(column: 2, row: 0))
        XCTAssertEqual(result, .number(13))
    }

    func testNestedFunctions() {
        sheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "3")
        sheet.setRawInput(at: CellAddress(column: 0, row: 1), rawInput: "7")
        let result = engine.evaluate(formula: "SUM(A1,MAX(A1,A2))", in: sheet, at: CellAddress(column: 1, row: 0))
        XCTAssertEqual(result, .number(10))
    }

    // MARK: - Comparisons

    func testComparisons() {
        let eq = engine.evaluate(formula: "1=1", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(eq, .boolean(true))

        let neq = engine.evaluate(formula: "1=2", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(neq, .boolean(false))

        let lt = engine.evaluate(formula: "1<2", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(lt, .boolean(true))

        let gt = engine.evaluate(formula: "2>1", in: sheet, at: CellAddress(column: 0, row: 0))
        XCTAssertEqual(gt, .boolean(true))
    }

    // MARK: - Recalculation

    func testRecalculation() {
        // A1 = 10
        sheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "10")
        // B1 = =A1*2
        sheet.setRawInput(at: CellAddress(column: 1, row: 0), rawInput: "=A1*2")
        // C1 = =B1+5
        sheet.setRawInput(at: CellAddress(column: 2, row: 0), rawInput: "=B1+5")

        engine.recalculate(worksheet: sheet)

        XCTAssertEqual(sheet.cellValue(at: CellAddress(column: 1, row: 0)), .number(20))
        XCTAssertEqual(sheet.cellValue(at: CellAddress(column: 2, row: 0)), .number(25))
    }

    func testRecalculationAfterChange() {
        sheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "10")
        sheet.setRawInput(at: CellAddress(column: 1, row: 0), rawInput: "=A1*2")

        engine.recalculate(worksheet: sheet)
        XCTAssertEqual(sheet.cellValue(at: CellAddress(column: 1, row: 0)), .number(20))

        // Change A1
        sheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "5")
        engine.recalculate(worksheet: sheet)
        XCTAssertEqual(sheet.cellValue(at: CellAddress(column: 1, row: 0)), .number(10))
    }

    // MARK: - Dependency Graph

    func testDependencyGraph() {
        let graph = DependencyGraph()
        let a1 = CellAddress(column: 0, row: 0)
        let b1 = CellAddress(column: 1, row: 0)
        let c1 = CellAddress(column: 2, row: 0)

        graph.addDependencies(cell: b1, dependsOn: [a1])
        graph.addDependencies(cell: c1, dependsOn: [b1])

        let deps = graph.dependents(of: a1)
        XCTAssertTrue(deps.contains(b1))
        XCTAssertTrue(deps.contains(c1))
    }

    func testCycleDetection() {
        let graph = DependencyGraph()
        let a1 = CellAddress(column: 0, row: 0)
        let b1 = CellAddress(column: 1, row: 0)

        graph.addDependencies(cell: a1, dependsOn: [b1])
        graph.addDependencies(cell: b1, dependsOn: [a1])

        XCTAssertTrue(graph.detectCycle(from: a1))
    }
}
