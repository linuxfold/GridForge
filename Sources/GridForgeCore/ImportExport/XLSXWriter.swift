import Foundation
import ZIPFoundation

/// Writes a Workbook to a valid .xlsx file (ZIP archive with XML parts)
public class XLSXWriter {

    /// Write a workbook to an XLSX file at the given URL
    public static func write(_ workbook: Workbook, to url: URL) throws {
        // Remove existing file if present
        if FileManager.default.fileExists(atPath: url.path) {
            try FileManager.default.removeItem(at: url)
        }

        let archive: Archive
        do {
            archive = try Archive(url: url, accessMode: .create)
        } catch {
            throw XLSXWriterError.cannotCreateArchive(url)
        }

        // Build shared string table from all string cells across all sheets
        var sharedStringSet: [String: Int] = [:]
        var sharedStrings: [String] = []
        for sheet in workbook.sheets {
            for (_, cell) in sheet.cells {
                if case .string(let s) = cell.value, !cell.isFormula {
                    if sharedStringSet[s] == nil {
                        sharedStringSet[s] = sharedStrings.count
                        sharedStrings.append(s)
                    }
                }
            }
        }

        // 1. [Content_Types].xml
        try addEntry(to: archive, path: "[Content_Types].xml",
                     content: contentTypesXML(sheetCount: workbook.sheets.count,
                                              hasSharedStrings: !sharedStrings.isEmpty))

        // 2. _rels/.rels
        try addEntry(to: archive, path: "_rels/.rels", content: topLevelRelsXML())

        // 3. xl/workbook.xml
        try addEntry(to: archive, path: "xl/workbook.xml",
                     content: workbookXML(sheets: workbook.sheets))

        // 4. xl/_rels/workbook.xml.rels
        try addEntry(to: archive, path: "xl/_rels/workbook.xml.rels",
                     content: workbookRelsXML(sheetCount: workbook.sheets.count,
                                              hasSharedStrings: !sharedStrings.isEmpty))

        // 5. xl/styles.xml (minimal)
        try addEntry(to: archive, path: "xl/styles.xml", content: minimalStylesXML())

        // 6. xl/sharedStrings.xml
        if !sharedStrings.isEmpty {
            try addEntry(to: archive, path: "xl/sharedStrings.xml",
                         content: sharedStringsXML(sharedStrings))
        }

        // 7. Worksheets
        for (index, sheet) in workbook.sheets.enumerated() {
            let sheetXML = worksheetXML(sheet: sheet, sharedStringSet: sharedStringSet)
            try addEntry(to: archive, path: "xl/worksheets/sheet\(index + 1).xml",
                         content: sheetXML)
        }
    }

    // MARK: - Archive Helpers

    private static func addEntry(to archive: Archive, path: String, content: String) throws {
        let data = Data(content.utf8)
        let uncompressedSize = Int64(data.count)
        var offset = 0
        try archive.addEntry(
            with: path,
            type: .file,
            uncompressedSize: uncompressedSize
        ) { (position: Int64, size: Int) -> Data in
            let start = offset
            let end = min(start + size, data.count)
            offset = end
            if start >= data.count {
                return Data()
            }
            return data[start..<end]
        }
    }

    // MARK: - XML Generation

    private static func contentTypesXML(sheetCount: Int, hasSharedStrings: Bool) -> String {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
        """

        if hasSharedStrings {
            xml += """

              <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
            """
        }

        for i in 1...sheetCount {
            xml += """

              <Override PartName="/xl/worksheets/sheet\(i).xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
            """
        }

        xml += "\n</Types>"
        return xml
    }

    private static func topLevelRelsXML() -> String {
        return """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """
    }

    private static func workbookXML(sheets: [Worksheet]) -> String {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>
        """

        for (index, sheet) in sheets.enumerated() {
            let escapedName = escapeXML(sheet.name)
            xml += """

            <sheet name="\(escapedName)" sheetId="\(index + 1)" r:id="rId\(index + 1)"/>
            """
        }

        xml += """

          </sheets>
        </workbook>
        """
        return xml
    }

    private static func workbookRelsXML(sheetCount: Int, hasSharedStrings: Bool) -> String {
        var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        xml += "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"

        for i in 1...sheetCount {
            xml += "\n  <Relationship Id=\"rId\(i)\""
            xml += " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\""
            xml += " Target=\"worksheets/sheet\(i).xml\"/>"
        }

        let stylesId = sheetCount + 1
        xml += "\n  <Relationship Id=\"rId\(stylesId)\""
        xml += " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\""
        xml += " Target=\"styles.xml\"/>"

        if hasSharedStrings {
            let ssId = sheetCount + 2
            xml += "\n  <Relationship Id=\"rId\(ssId)\""
            xml += " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\""
            xml += " Target=\"sharedStrings.xml\"/>"
        }

        xml += "\n</Relationships>"
        return xml
    }

    private static func minimalStylesXML() -> String {
        return """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="1">
            <font>
              <sz val="11"/>
              <name val="Calibri"/>
            </font>
          </fonts>
          <fills count="2">
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="gray125"/></fill>
          </fills>
          <borders count="1">
            <border>
              <left/><right/><top/><bottom/><diagonal/>
            </border>
          </borders>
          <cellStyleXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
          </cellStyleXfs>
          <cellXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
          </cellXfs>
        </styleSheet>
        """
    }

    private static func sharedStringsXML(_ strings: [String]) -> String {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="\(strings.count)" uniqueCount="\(strings.count)">
        """

        for s in strings {
            xml += "\n  <si><t>\(escapeXML(s))</t></si>"
        }

        xml += "\n</sst>"
        return xml
    }

    private static func worksheetXML(sheet: Worksheet, sharedStringSet: [String: Int]) -> String {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
        """

        // Gather all cells and organize by row
        var rowMap: [Int: [(CellAddress, Cell)]] = [:]
        for (address, cell) in sheet.cells {
            if cell.isEmpty { continue }
            rowMap[address.row, default: []].append((address, cell))
        }

        // Sort rows
        let sortedRowKeys = rowMap.keys.sorted()

        for rowIndex in sortedRowKeys {
            guard var rowCells = rowMap[rowIndex] else { continue }
            rowCells.sort { $0.0.column < $1.0.column }

            xml += "\n    <row r=\"\(rowIndex + 1)\">"

            for (address, cell) in rowCells {
                let ref = address.displayString

                if cell.isFormula {
                    // Formula cell
                    let formula = cell.formulaExpression ?? ""
                    let escapedFormula = escapeXML(formula)

                    switch cell.value {
                    case .string(let s):
                        xml += "\n      <c r=\"\(ref)\" t=\"str\"><f>\(escapedFormula)</f><v>\(escapeXML(s))</v></c>"
                    case .boolean(let b):
                        xml += "\n      <c r=\"\(ref)\" t=\"b\"><f>\(escapedFormula)</f><v>\(b ? "1" : "0")</v></c>"
                    case .number(let n):
                        xml += "\n      <c r=\"\(ref)\"><f>\(escapedFormula)</f><v>\(formatNumber(n))</v></c>"
                    case .error(let e):
                        xml += "\n      <c r=\"\(ref)\" t=\"e\"><f>\(escapedFormula)</f><v>\(escapeXML(e.rawValue))</v></c>"
                    default:
                        xml += "\n      <c r=\"\(ref)\"><f>\(escapedFormula)</f></c>"
                    }
                } else {
                    switch cell.value {
                    case .string(let s):
                        if let idx = sharedStringSet[s] {
                            xml += "\n      <c r=\"\(ref)\" t=\"s\"><v>\(idx)</v></c>"
                        } else {
                            xml += "\n      <c r=\"\(ref)\" t=\"str\"><v>\(escapeXML(s))</v></c>"
                        }
                    case .number(let n):
                        xml += "\n      <c r=\"\(ref)\"><v>\(formatNumber(n))</v></c>"
                    case .boolean(let b):
                        xml += "\n      <c r=\"\(ref)\" t=\"b\"><v>\(b ? "1" : "0")</v></c>"
                    case .error(let e):
                        xml += "\n      <c r=\"\(ref)\" t=\"e\"><v>\(escapeXML(e.rawValue))</v></c>"
                    case .date(let d):
                        // Excel date serial number (days since 1900-01-01, with the 1900 bug)
                        let serial = excelDateSerial(from: d)
                        xml += "\n      <c r=\"\(ref)\"><v>\(formatNumber(serial))</v></c>"
                    case .empty:
                        break
                    }
                }
            }

            xml += "\n    </row>"
        }

        xml += """

          </sheetData>
        </worksheet>
        """
        return xml
    }

    // MARK: - Utility

    private static func escapeXML(_ string: String) -> String {
        var result = string
        result = result.replacingOccurrences(of: "&", with: "&amp;")
        result = result.replacingOccurrences(of: "<", with: "&lt;")
        result = result.replacingOccurrences(of: ">", with: "&gt;")
        result = result.replacingOccurrences(of: "\"", with: "&quot;")
        result = result.replacingOccurrences(of: "'", with: "&apos;")
        return result
    }

    private static func formatNumber(_ n: Double) -> String {
        if n == n.rounded(.towardZero) && abs(n) < 1e15 && n == Double(Int64(n)) {
            return String(Int64(n))
        }
        return String(n)
    }

    /// Convert a Swift Date to an Excel date serial number
    /// Excel uses a serial date system where 1 = 1900-01-01
    /// Note: Excel incorrectly treats 1900 as a leap year (the "1900 bug")
    private static func excelDateSerial(from date: Date) -> Double {
        let calendar = Calendar(identifier: .gregorian)
        // Excel epoch: 1899-12-30 (to account for the 1900 leap year bug)
        var components = DateComponents()
        components.year = 1899
        components.month = 12
        components.day = 30
        components.hour = 0
        components.minute = 0
        components.second = 0
        guard let epoch = calendar.date(from: components) else {
            return 0
        }
        let interval = date.timeIntervalSince(epoch)
        return interval / 86400.0
    }
}

// MARK: - Writer Errors

public enum XLSXWriterError: Error, LocalizedError {
    case cannotCreateArchive(URL)
    case writeFailed(String)

    public var errorDescription: String? {
        switch self {
        case .cannotCreateArchive(let url):
            return "Cannot create XLSX archive at \(url.path)"
        case .writeFailed(let detail):
            return "XLSX write failed: \(detail)"
        }
    }
}
