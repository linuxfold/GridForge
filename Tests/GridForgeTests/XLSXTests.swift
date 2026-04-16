import XCTest
@testable import GridForgeCore

final class XLSXTests: XCTestCase {

    var tempDir: URL!

    override func setUp() {
        super.setUp()
        tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("GridForgeTests_\(UUID().uuidString)")
        try? FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
    }

    override func tearDown() {
        try? FileManager.default.removeItem(at: tempDir)
        super.tearDown()
    }

    // MARK: - Export

    func testExportCreatesFile() throws {
        let wb = Workbook()
        wb.activeSheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "Hello")

        let url = tempDir.appendingPathComponent("test.xlsx")
        try XLSXWriter.write(wb, to: url)

        XCTAssertTrue(FileManager.default.fileExists(atPath: url.path))
    }

    func testExportMultipleSheets() throws {
        let wb = Workbook()
        wb.activeSheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "Sheet1Data")
        wb.addSheet(name: "Second")
        wb.sheets[1].setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "Sheet2Data")

        let url = tempDir.appendingPathComponent("multi.xlsx")
        try XLSXWriter.write(wb, to: url)

        XCTAssertTrue(FileManager.default.fileExists(atPath: url.path))
    }

    // MARK: - Roundtrip

    func testRoundtripStrings() throws {
        let wb = Workbook()
        wb.activeSheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "Hello")
        wb.activeSheet.setRawInput(at: CellAddress(column: 1, row: 0), rawInput: "World")

        let url = tempDir.appendingPathComponent("strings.xlsx")
        try XLSXWriter.write(wb, to: url)

        let imported = try XLSXReader.read(from: url)
        XCTAssertEqual(imported.sheets.count, 1)
        XCTAssertEqual(imported.sheets[0].cellValue(at: CellAddress(column: 0, row: 0)), .string("Hello"))
        XCTAssertEqual(imported.sheets[0].cellValue(at: CellAddress(column: 1, row: 0)), .string("World"))
    }

    func testRoundtripNumbers() throws {
        let wb = Workbook()
        wb.activeSheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "42")
        wb.activeSheet.setRawInput(at: CellAddress(column: 1, row: 0), rawInput: "3.14")

        let url = tempDir.appendingPathComponent("numbers.xlsx")
        try XLSXWriter.write(wb, to: url)

        let imported = try XLSXReader.read(from: url)
        XCTAssertEqual(imported.sheets[0].cellValue(at: CellAddress(column: 0, row: 0)), .number(42))
        XCTAssertEqual(imported.sheets[0].cellValue(at: CellAddress(column: 1, row: 0)), .number(3.14))
    }

    func testRoundtripBooleans() throws {
        let wb = Workbook()
        wb.activeSheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "TRUE")
        wb.activeSheet.setRawInput(at: CellAddress(column: 1, row: 0), rawInput: "FALSE")

        let url = tempDir.appendingPathComponent("booleans.xlsx")
        try XLSXWriter.write(wb, to: url)

        let imported = try XLSXReader.read(from: url)
        XCTAssertEqual(imported.sheets[0].cellValue(at: CellAddress(column: 0, row: 0)), .boolean(true))
        XCTAssertEqual(imported.sheets[0].cellValue(at: CellAddress(column: 1, row: 0)), .boolean(false))
    }

    func testRoundtripMultipleSheets() throws {
        let wb = Workbook()
        wb.activeSheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "First")
        wb.addSheet(name: "Second")
        wb.sheets[1].setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "Second")
        wb.addSheet(name: "Third")
        wb.sheets[2].setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "Third")

        let url = tempDir.appendingPathComponent("sheets.xlsx")
        try XLSXWriter.write(wb, to: url)

        let imported = try XLSXReader.read(from: url)
        XCTAssertEqual(imported.sheets.count, 3)
        XCTAssertEqual(imported.sheets[0].name, "Sheet1")
        XCTAssertEqual(imported.sheets[1].name, "Second")
        XCTAssertEqual(imported.sheets[2].name, "Third")
        XCTAssertEqual(imported.sheets[0].cellValue(at: CellAddress(column: 0, row: 0)), .string("First"))
        XCTAssertEqual(imported.sheets[1].cellValue(at: CellAddress(column: 0, row: 0)), .string("Second"))
        XCTAssertEqual(imported.sheets[2].cellValue(at: CellAddress(column: 0, row: 0)), .string("Third"))
    }

    func testRoundtripFormulas() throws {
        let wb = Workbook()
        wb.activeSheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "10")
        wb.activeSheet.setRawInput(at: CellAddress(column: 0, row: 1), rawInput: "20")
        let formulaCell = wb.activeSheet.setRawInput(at: CellAddress(column: 0, row: 2), rawInput: "=SUM(A1:A2)")
        formulaCell.value = .number(30) // simulate formula result

        let url = tempDir.appendingPathComponent("formulas.xlsx")
        try XLSXWriter.write(wb, to: url)

        let imported = try XLSXReader.read(from: url)
        let cell = imported.sheets[0].cell(at: CellAddress(column: 0, row: 2))
        XCTAssertNotNil(cell)
        XCTAssertTrue(cell?.isFormula ?? false)
        XCTAssertEqual(cell?.formulaExpression, "SUM(A1:A2)")
    }

    func testRoundtripMixedContent() throws {
        let wb = Workbook()
        let s = wb.activeSheet
        s.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "Name")
        s.setRawInput(at: CellAddress(column: 1, row: 0), rawInput: "Value")
        s.setRawInput(at: CellAddress(column: 0, row: 1), rawInput: "Alpha")
        s.setRawInput(at: CellAddress(column: 1, row: 1), rawInput: "100")
        s.setRawInput(at: CellAddress(column: 0, row: 2), rawInput: "Beta")
        s.setRawInput(at: CellAddress(column: 1, row: 2), rawInput: "200")
        s.setRawInput(at: CellAddress(column: 0, row: 3), rawInput: "Total")
        let totalCell = s.setRawInput(at: CellAddress(column: 1, row: 3), rawInput: "=SUM(B2:B3)")
        totalCell.value = .number(300)

        let url = tempDir.appendingPathComponent("mixed.xlsx")
        try XLSXWriter.write(wb, to: url)

        let imported = try XLSXReader.read(from: url)
        let is_ = imported.sheets[0]
        XCTAssertEqual(is_.cellValue(at: CellAddress(column: 0, row: 0)), .string("Name"))
        XCTAssertEqual(is_.cellValue(at: CellAddress(column: 1, row: 0)), .string("Value"))
        XCTAssertEqual(is_.cellValue(at: CellAddress(column: 0, row: 1)), .string("Alpha"))
        XCTAssertEqual(is_.cellValue(at: CellAddress(column: 1, row: 1)), .number(100))
        XCTAssertEqual(is_.cellValue(at: CellAddress(column: 0, row: 2)), .string("Beta"))
        XCTAssertEqual(is_.cellValue(at: CellAddress(column: 1, row: 2)), .number(200))

        let total = is_.cell(at: CellAddress(column: 1, row: 3))
        XCTAssertTrue(total?.isFormula ?? false)
    }

    func testRoundtripLargeSheet() throws {
        let wb = Workbook()
        let s = wb.activeSheet
        for row in 0..<100 {
            for col in 0..<10 {
                s.setRawInput(at: CellAddress(column: col, row: row), rawInput: "\(row * 10 + col)")
            }
        }

        let url = tempDir.appendingPathComponent("large.xlsx")
        try XLSXWriter.write(wb, to: url)

        let imported = try XLSXReader.read(from: url)
        XCTAssertEqual(imported.sheets[0].cellValue(at: CellAddress(column: 0, row: 0)), .number(0))
        XCTAssertEqual(imported.sheets[0].cellValue(at: CellAddress(column: 9, row: 99)), .number(999))
        XCTAssertEqual(imported.sheets[0].cells.count, 1000)
    }

    func testRoundtripSpecialCharacters() throws {
        let wb = Workbook()
        wb.activeSheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "Hello & World")
        wb.activeSheet.setRawInput(at: CellAddress(column: 0, row: 1), rawInput: "<tag>")
        wb.activeSheet.setRawInput(at: CellAddress(column: 0, row: 2), rawInput: "It's \"quoted\"")

        let url = tempDir.appendingPathComponent("special.xlsx")
        try XLSXWriter.write(wb, to: url)

        let imported = try XLSXReader.read(from: url)
        XCTAssertEqual(imported.sheets[0].cellValue(at: CellAddress(column: 0, row: 0)), .string("Hello & World"))
        XCTAssertEqual(imported.sheets[0].cellValue(at: CellAddress(column: 0, row: 1)), .string("<tag>"))
        XCTAssertEqual(imported.sheets[0].cellValue(at: CellAddress(column: 0, row: 2)), .string("It's \"quoted\""))
    }

    // MARK: - Error Cases

    func testImportNonexistentFile() {
        let url = tempDir.appendingPathComponent("nonexistent.xlsx")
        XCTAssertThrowsError(try XLSXReader.read(from: url))
    }

    func testExportEmptyWorkbook() throws {
        let wb = Workbook()
        let url = tempDir.appendingPathComponent("empty.xlsx")
        try XLSXWriter.write(wb, to: url)

        let imported = try XLSXReader.read(from: url)
        XCTAssertEqual(imported.sheets.count, 1)
        XCTAssertTrue(imported.sheets[0].cells.isEmpty)
    }
}
