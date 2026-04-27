import XCTest
@testable import GridForgeCore

final class ModelTests: XCTestCase {

    // MARK: - CellAddress

    func testCellAddressColumnLetter() {
        XCTAssertEqual(CellAddress(column: 0, row: 0).columnLetter, "A")
        XCTAssertEqual(CellAddress(column: 25, row: 0).columnLetter, "Z")
        XCTAssertEqual(CellAddress(column: 26, row: 0).columnLetter, "AA")
        XCTAssertEqual(CellAddress(column: 27, row: 0).columnLetter, "AB")
        XCTAssertEqual(CellAddress(column: 51, row: 0).columnLetter, "AZ")
        XCTAssertEqual(CellAddress(column: 52, row: 0).columnLetter, "BA")
    }

    func testCellAddressDisplayString() {
        XCTAssertEqual(CellAddress(column: 0, row: 0).displayString, "A1")
        XCTAssertEqual(CellAddress(column: 2, row: 4).displayString, "C5")
        XCTAssertEqual(CellAddress(column: 26, row: 99).displayString, "AA100")
    }

    func testCellAddressParse() {
        let a1 = CellAddress.parse("A1")
        XCTAssertEqual(a1, CellAddress(column: 0, row: 0))

        let c5 = CellAddress.parse("C5")
        XCTAssertEqual(c5, CellAddress(column: 2, row: 4))

        let aa100 = CellAddress.parse("AA100")
        XCTAssertEqual(aa100, CellAddress(column: 26, row: 99))

        XCTAssertNil(CellAddress.parse(""))
        XCTAssertNil(CellAddress.parse("123"))
        XCTAssertNil(CellAddress.parse("ABC"))
        XCTAssertNil(CellAddress.parse("A0"))
    }

    func testCellAddressRoundTrip() {
        for col in 0..<100 {
            for row in 0..<10 {
                let addr = CellAddress(column: col, row: row)
                let parsed = CellAddress.parse(addr.displayString)
                XCTAssertEqual(parsed, addr, "Failed for \(addr.displayString)")
            }
        }
    }

    func testCellReferenceParsePreservesAbsoluteMarkers() {
        let ref = CellReference.parse("$B$12")
        XCTAssertEqual(ref?.address, CellAddress(column: 1, row: 11))
        XCTAssertEqual(ref?.columnAbsolute, true)
        XCTAssertEqual(ref?.rowAbsolute, true)
        XCTAssertEqual(ref?.displayString, "$B$12")

        let mixed = CellReference.parse("C$3")
        XCTAssertEqual(mixed?.address, CellAddress(column: 2, row: 2))
        XCTAssertEqual(mixed?.columnAbsolute, false)
        XCTAssertEqual(mixed?.rowAbsolute, true)
    }

    // MARK: - CellRange

    func testCellRangeNormalization() {
        let range = CellRange(
            start: CellAddress(column: 3, row: 5),
            end: CellAddress(column: 1, row: 2)
        )
        XCTAssertEqual(range.start, CellAddress(column: 1, row: 2))
        XCTAssertEqual(range.end, CellAddress(column: 3, row: 5))
    }

    func testCellRangeContains() {
        let range = CellRange(
            start: CellAddress(column: 1, row: 1),
            end: CellAddress(column: 3, row: 3)
        )
        XCTAssertTrue(range.contains(CellAddress(column: 2, row: 2)))
        XCTAssertTrue(range.contains(CellAddress(column: 1, row: 1)))
        XCTAssertTrue(range.contains(CellAddress(column: 3, row: 3)))
        XCTAssertFalse(range.contains(CellAddress(column: 0, row: 0)))
        XCTAssertFalse(range.contains(CellAddress(column: 4, row: 2)))
    }

    func testCellRangeAllAddresses() {
        let range = CellRange(
            start: CellAddress(column: 0, row: 0),
            end: CellAddress(column: 1, row: 1)
        )
        let all = range.allAddresses
        XCTAssertEqual(all.count, 4)
        XCTAssertTrue(all.contains(CellAddress(column: 0, row: 0)))
        XCTAssertTrue(all.contains(CellAddress(column: 1, row: 0)))
        XCTAssertTrue(all.contains(CellAddress(column: 0, row: 1)))
        XCTAssertTrue(all.contains(CellAddress(column: 1, row: 1)))
    }

    // MARK: - CellValue

    func testCellValueDisplayString() {
        XCTAssertEqual(CellValue.empty.displayString, "")
        XCTAssertEqual(CellValue.string("hello").displayString, "hello")
        XCTAssertEqual(CellValue.number(42).displayString, "42")
        XCTAssertEqual(CellValue.number(3.14).displayString, "3.14")
        XCTAssertEqual(CellValue.boolean(true).displayString, "TRUE")
        XCTAssertEqual(CellValue.boolean(false).displayString, "FALSE")
        XCTAssertEqual(CellValue.error(.divZero).displayString, "#DIV/0!")
    }

    func testCellValueNumericValue() {
        XCTAssertEqual(CellValue.number(42).numericValue, 42)
        XCTAssertEqual(CellValue.boolean(true).numericValue, 1)
        XCTAssertEqual(CellValue.boolean(false).numericValue, 0)
        XCTAssertNil(CellValue.string("hello").numericValue)
        XCTAssertEqual(CellValue.empty.numericValue, 0)  // empty is 0 in numeric context (Excel convention)
    }

    // MARK: - Cell

    func testCellFormula() {
        let cell = Cell(rawInput: "=SUM(A1:A10)")
        XCTAssertTrue(cell.isFormula)
        XCTAssertEqual(cell.formulaExpression, "SUM(A1:A10)")

        let valueCell = Cell(rawInput: "42")
        XCTAssertFalse(valueCell.isFormula)
        XCTAssertNil(valueCell.formulaExpression)
    }

    func testCellCopy() {
        let original = Cell(rawInput: "hello", value: .string("hello"))
        original.formatting.bold = true
        let copy = original.copy()
        XCTAssertEqual(copy.rawInput, "hello")
        XCTAssertEqual(copy.value, .string("hello"))
        XCTAssertTrue(copy.formatting.bold)
    }

    // MARK: - Worksheet

    func testWorksheetSetAndGet() {
        let sheet = Worksheet(name: "Test")
        let addr = CellAddress(column: 0, row: 0)
        sheet.setRawInput(at: addr, rawInput: "42")
        XCTAssertEqual(sheet.cellValue(at: addr), .number(42))

        sheet.setRawInput(at: addr, rawInput: "hello")
        XCTAssertEqual(sheet.cellValue(at: addr), .string("hello"))

        sheet.setRawInput(at: addr, rawInput: "TRUE")
        XCTAssertEqual(sheet.cellValue(at: addr), .boolean(true))
    }

    func testWorksheetClear() {
        let sheet = Worksheet(name: "Test")
        let addr = CellAddress(column: 0, row: 0)
        sheet.setRawInput(at: addr, rawInput: "42")
        XCTAssertNotNil(sheet.cell(at: addr))

        sheet.clearCell(at: addr)
        XCTAssertNil(sheet.cell(at: addr))
        XCTAssertEqual(sheet.cellValue(at: addr), .empty)
    }

    func testWorksheetInsertRow() {
        let sheet = Worksheet(name: "Test")
        sheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "A")
        sheet.setRawInput(at: CellAddress(column: 0, row: 1), rawInput: "B")
        sheet.setRawInput(at: CellAddress(column: 0, row: 2), rawInput: "C")

        sheet.insertRow(at: 1)

        XCTAssertEqual(sheet.cellValue(at: CellAddress(column: 0, row: 0)), .string("A"))
        XCTAssertEqual(sheet.cellValue(at: CellAddress(column: 0, row: 1)), .empty)
        XCTAssertEqual(sheet.cellValue(at: CellAddress(column: 0, row: 2)), .string("B"))
        XCTAssertEqual(sheet.cellValue(at: CellAddress(column: 0, row: 3)), .string("C"))
    }

    func testWorksheetDeleteRow() {
        let sheet = Worksheet(name: "Test")
        sheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "A")
        sheet.setRawInput(at: CellAddress(column: 0, row: 1), rawInput: "B")
        sheet.setRawInput(at: CellAddress(column: 0, row: 2), rawInput: "C")

        sheet.deleteRow(at: 1)

        XCTAssertEqual(sheet.cellValue(at: CellAddress(column: 0, row: 0)), .string("A"))
        XCTAssertEqual(sheet.cellValue(at: CellAddress(column: 0, row: 1)), .string("C"))
    }

    func testWorksheetUsedRange() {
        let sheet = Worksheet(name: "Test")
        XCTAssertNil(sheet.usedRange)

        sheet.setRawInput(at: CellAddress(column: 1, row: 2), rawInput: "X")
        sheet.setRawInput(at: CellAddress(column: 3, row: 5), rawInput: "Y")

        let range = sheet.usedRange
        XCTAssertEqual(range?.start, CellAddress(column: 1, row: 2))
        XCTAssertEqual(range?.end, CellAddress(column: 3, row: 5))
    }

    // MARK: - Workbook

    func testWorkbookSheetManagement() {
        let wb = Workbook()
        XCTAssertEqual(wb.sheets.count, 1)
        XCTAssertEqual(wb.sheets[0].name, "Sheet1")

        wb.addSheet(name: "Data")
        XCTAssertEqual(wb.sheets.count, 2)
        XCTAssertEqual(wb.sheets[1].name, "Data")

        wb.deleteSheet(at: 0)
        XCTAssertEqual(wb.sheets.count, 1)
        XCTAssertEqual(wb.sheets[0].name, "Data")

        // Can't delete last sheet
        wb.deleteSheet(at: 0)
        XCTAssertEqual(wb.sheets.count, 1)
    }

    func testWorkbookDuplicate() {
        let wb = Workbook()
        wb.activeSheet.setRawInput(at: CellAddress(column: 0, row: 0), rawInput: "hello")
        wb.duplicateSheet(at: 0)

        XCTAssertEqual(wb.sheets.count, 2)
        XCTAssertEqual(wb.sheets[1].cellValue(at: CellAddress(column: 0, row: 0)), .string("hello"))
    }

    func testWorkbookRename() {
        let wb = Workbook()
        wb.renameSheet(at: 0, to: "MySheet")
        XCTAssertEqual(wb.sheets[0].name, "MySheet")
    }

    // MARK: - Selection

    func testSelectionMove() {
        var sel = SelectionState()
        XCTAssertEqual(sel.activeCell, CellAddress(column: 0, row: 0))

        sel.moveActiveCell(direction: .right, extend: false, maxColumn: 25, maxRow: 99)
        XCTAssertEqual(sel.activeCell, CellAddress(column: 1, row: 0))

        sel.moveActiveCell(direction: .down, extend: false, maxColumn: 25, maxRow: 99)
        XCTAssertEqual(sel.activeCell, CellAddress(column: 1, row: 1))

        sel.moveActiveCell(direction: .left, extend: false, maxColumn: 25, maxRow: 99)
        XCTAssertEqual(sel.activeCell, CellAddress(column: 0, row: 1))

        sel.moveActiveCell(direction: .up, extend: false, maxColumn: 25, maxRow: 99)
        XCTAssertEqual(sel.activeCell, CellAddress(column: 0, row: 0))
    }

    func testSelectionBoundary() {
        var sel = SelectionState()
        sel.moveActiveCell(direction: .left, extend: false, maxColumn: 25, maxRow: 99)
        XCTAssertEqual(sel.activeCell.column, 0)

        sel.moveActiveCell(direction: .up, extend: false, maxColumn: 25, maxRow: 99)
        XCTAssertEqual(sel.activeCell.row, 0)
    }

    func testSelectionExtend() {
        var sel = SelectionState()
        sel.moveActiveCell(direction: .right, extend: true, maxColumn: 25, maxRow: 99)
        sel.moveActiveCell(direction: .down, extend: true, maxColumn: 25, maxRow: 99)
        XCTAssertEqual(sel.selectedRange.start, CellAddress(column: 0, row: 0))
        XCTAssertEqual(sel.selectedRange.end, CellAddress(column: 1, row: 1))
    }
}
