import Foundation
import ZIPFoundation

// MARK: - Error Types

/// Errors that can occur during XLSX reading
public enum XLSXError: Error, LocalizedError {
    case cannotOpenArchive(URL)
    case missingEntry(String)
    case corruptedData(String)
    case xmlParsingFailed(String)
    case invalidCellReference(String)
    case unsupportedFormat(String)

    public var errorDescription: String? {
        switch self {
        case .cannotOpenArchive(let url):
            return "Cannot open XLSX archive at \(url.path)"
        case .missingEntry(let path):
            return "Missing required entry in XLSX: \(path)"
        case .corruptedData(let detail):
            return "Corrupted data in XLSX: \(detail)"
        case .xmlParsingFailed(let detail):
            return "XML parsing failed: \(detail)"
        case .invalidCellReference(let ref):
            return "Invalid cell reference: \(ref)"
        case .unsupportedFormat(let detail):
            return "Unsupported format: \(detail)"
        }
    }
}

// MARK: - XLSXReader

/// Reads .xlsx files (ZIP archives containing XML) into Workbook models
public class XLSXReader {

    /// Read an XLSX file and produce a Workbook
    public static func read(from url: URL) throws -> Workbook {
        let archive: Archive
        do {
            archive = try Archive(url: url, accessMode: .read)
        } catch {
            throw XLSXError.cannotOpenArchive(url)
        }
        let packageEntries = try captureEntries(from: archive)

        // 1. Parse shared strings
        let sharedStrings: [String]
        if let ssEntry = archive["xl/sharedStrings.xml"] {
            let ssData = try extractData(from: archive, entry: ssEntry)
            sharedStrings = try SharedStringsParser.parse(data: ssData)
        } else {
            sharedStrings = []
        }

        // 2. Parse workbook.xml to get sheet names and rIds
        let wbData = try extractEntry(from: archive, path: "xl/workbook.xml")
        let sheetInfos = try WorkbookXMLParser.parse(data: wbData)

        // 3. Parse xl/_rels/workbook.xml.rels to map rIds to file paths
        let relsData = try extractEntry(from: archive, path: "xl/_rels/workbook.xml.rels")
        let relationships = try RelationshipsParser.parse(data: relsData)

        // 4. For each sheet, parse the worksheet XML
        var worksheets: [Worksheet] = []
        var sheetPartPaths: [String] = []
        for info in sheetInfos {
            guard let rel = relationships[info.rId] else {
                throw XLSXError.corruptedData("No relationship found for rId=\(info.rId)")
            }
            // rel.target is like "worksheets/sheet1.xml"; prefix with "xl/"
            let sheetPath: String
            if rel.hasPrefix("/") {
                // Absolute path within archive (strip leading /)
                sheetPath = String(rel.dropFirst())
            } else {
                sheetPath = "xl/" + rel
            }
            sheetPartPaths.append(sheetPath)

            let sheetData = try extractEntry(from: archive, path: sheetPath)
            let cells = try SheetParser.parse(data: sheetData, sharedStrings: sharedStrings)

            let worksheet = Worksheet(name: info.name)
            for (address, cell) in cells {
                worksheet.cells[address] = cell
            }
            worksheets.append(worksheet)
        }

        if worksheets.isEmpty {
            worksheets = [Worksheet(name: "Sheet1")]
        }

        let workbook = Workbook(
            sheets: worksheets,
            sourceXLSXPackage: XLSXPackage(entries: packageEntries, sheetPartPaths: sheetPartPaths)
        )
        return workbook
    }

    // MARK: - Helpers

    private static func extractData(from archive: Archive, entry: Entry) throws -> Data {
        var data = Data()
        _ = try archive.extract(entry) { chunk in
            data.append(chunk)
        }
        return data
    }

    private static func extractEntry(from archive: Archive, path: String) throws -> Data {
        guard let entry = archive[path] else {
            throw XLSXError.missingEntry(path)
        }
        return try extractData(from: archive, entry: entry)
    }

    private static func captureEntries(from archive: Archive) throws -> [String: Data] {
        var entries: [String: Data] = [:]
        for entry in archive {
            entries[entry.path] = try extractData(from: archive, entry: entry)
        }
        return entries
    }
}

// MARK: - Sheet Info

/// Minimal info parsed from workbook.xml
private struct SheetInfo {
    let name: String
    let sheetId: String
    let rId: String
}

// MARK: - SharedStringsParser

/// Parses xl/sharedStrings.xml into an array of strings
private class SharedStringsParser: NSObject, XMLParserDelegate {
    private var strings: [String] = []
    private var currentText = ""
    private var insideSI = false
    private var insideT = false

    static func parse(data: Data) throws -> [String] {
        let parser = XMLParser(data: data)
        let delegate = SharedStringsParser()
        parser.delegate = delegate
        if !parser.parse() {
            if let error = parser.parserError {
                throw XLSXError.xmlParsingFailed("sharedStrings.xml: \(error.localizedDescription)")
            }
            throw XLSXError.xmlParsingFailed("sharedStrings.xml: unknown error")
        }
        return delegate.strings
    }

    func parser(
        _ parser: XMLParser,
        didStartElement elementName: String,
        namespaceURI: String?,
        qualifiedName: String?,
        attributes: [String: String]
    ) {
        switch elementName {
        case "si":
            insideSI = true
            currentText = ""
        case "t":
            if insideSI {
                insideT = true
            }
        default:
            break
        }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if insideT {
            currentText += string
        }
    }

    func parser(
        _ parser: XMLParser,
        didEndElement elementName: String,
        namespaceURI: String?,
        qualifiedName: String?
    ) {
        switch elementName {
        case "si":
            strings.append(currentText)
            insideSI = false
            currentText = ""
        case "t":
            insideT = false
        default:
            break
        }
    }
}

// MARK: - WorkbookXMLParser

/// Parses xl/workbook.xml to extract sheet names and relationship IDs
private class WorkbookXMLParser: NSObject, XMLParserDelegate {
    private var sheets: [SheetInfo] = []

    static func parse(data: Data) throws -> [SheetInfo] {
        let parser = XMLParser(data: data)
        let delegate = WorkbookXMLParser()
        parser.delegate = delegate
        if !parser.parse() {
            if let error = parser.parserError {
                throw XLSXError.xmlParsingFailed("workbook.xml: \(error.localizedDescription)")
            }
            throw XLSXError.xmlParsingFailed("workbook.xml: unknown error")
        }
        return delegate.sheets
    }

    func parser(
        _ parser: XMLParser,
        didStartElement elementName: String,
        namespaceURI: String?,
        qualifiedName: String?,
        attributes: [String: String]
    ) {
        if elementName == "sheet" {
            let name = attributes["name"] ?? "Sheet"
            let sheetId = attributes["sheetId"] ?? "1"
            // The rId attribute may have a namespace prefix (r:id)
            let rId = attributes["r:id"] ?? attributes["rId"] ?? ""
            sheets.append(SheetInfo(name: name, sheetId: sheetId, rId: rId))
        }
    }
}

// MARK: - RelationshipsParser

/// Parses .rels XML files to map relationship IDs to target paths
private class RelationshipsParser: NSObject, XMLParserDelegate {
    private var relationships: [String: String] = [:]

    /// Returns a dictionary mapping rId -> target path
    static func parse(data: Data) throws -> [String: String] {
        let parser = XMLParser(data: data)
        let delegate = RelationshipsParser()
        parser.delegate = delegate
        if !parser.parse() {
            if let error = parser.parserError {
                throw XLSXError.xmlParsingFailed("rels: \(error.localizedDescription)")
            }
            throw XLSXError.xmlParsingFailed("rels: unknown error")
        }
        return delegate.relationships
    }

    func parser(
        _ parser: XMLParser,
        didStartElement elementName: String,
        namespaceURI: String?,
        qualifiedName: String?,
        attributes: [String: String]
    ) {
        if elementName == "Relationship" {
            if let id = attributes["Id"], let target = attributes["Target"] {
                relationships[id] = target
            }
        }
    }
}

// MARK: - SheetParser

/// Parses xl/worksheets/sheetN.xml into cells
private class SheetParser: NSObject, XMLParserDelegate {
    private let sharedStrings: [String]
    private var cells: [CellAddress: Cell] = [:]

    // Current parsing state
    private var currentCellRef: String?
    private var currentCellType: String?
    private var currentValueText = ""
    private var currentFormulaText = ""
    private var insideV = false
    private var insideF = false
    private var hasFormula = false

    init(sharedStrings: [String]) {
        self.sharedStrings = sharedStrings
        super.init()
    }

    static func parse(data: Data, sharedStrings: [String]) throws -> [CellAddress: Cell] {
        let parser = XMLParser(data: data)
        let delegate = SheetParser(sharedStrings: sharedStrings)
        parser.delegate = delegate
        if !parser.parse() {
            if let error = parser.parserError {
                throw XLSXError.xmlParsingFailed("worksheet: \(error.localizedDescription)")
            }
            throw XLSXError.xmlParsingFailed("worksheet: unknown error")
        }
        return delegate.cells
    }

    func parser(
        _ parser: XMLParser,
        didStartElement elementName: String,
        namespaceURI: String?,
        qualifiedName: String?,
        attributes: [String: String]
    ) {
        switch elementName {
        case "c":
            // Cell element: <c r="A1" t="s">
            currentCellRef = attributes["r"]
            currentCellType = attributes["t"]
            currentValueText = ""
            currentFormulaText = ""
            hasFormula = false
        case "v":
            insideV = true
            currentValueText = ""
        case "f":
            insideF = true
            hasFormula = true
            currentFormulaText = ""
        default:
            break
        }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if insideV {
            currentValueText += string
        } else if insideF {
            currentFormulaText += string
        }
    }

    func parser(
        _ parser: XMLParser,
        didEndElement elementName: String,
        namespaceURI: String?,
        qualifiedName: String?
    ) {
        switch elementName {
        case "v":
            insideV = false
        case "f":
            insideF = false
        case "c":
            finalizeCell()
        default:
            break
        }
    }

    private func finalizeCell() {
        guard let ref = currentCellRef,
              let address = CellAddress.parse(ref) else {
            currentCellRef = nil
            return
        }

        let cellType = currentCellType ?? ""
        let rawValue = currentValueText.trimmingCharacters(in: .whitespacesAndNewlines)
        let formulaText = currentFormulaText.trimmingCharacters(in: .whitespacesAndNewlines)

        var cellValue: CellValue = .empty
        var rawInput: String = ""

        if hasFormula && !formulaText.isEmpty {
            // Formula cell
            rawInput = "=" + formulaText
            // The <v> element contains the cached result
            cellValue = resolveValue(rawValue: rawValue, type: cellType)
        } else {
            switch cellType {
            case "s":
                // Shared string
                if let idx = Int(rawValue), idx >= 0, idx < sharedStrings.count {
                    let str = sharedStrings[idx]
                    cellValue = .string(str)
                    rawInput = str
                }
            case "b":
                // Boolean
                let boolVal = rawValue == "1"
                cellValue = .boolean(boolVal)
                rawInput = boolVal ? "TRUE" : "FALSE"
            case "str":
                // Inline string (formula result stored as string)
                cellValue = .string(rawValue)
                rawInput = rawValue
            case "e":
                // Error
                cellValue = .error(mapErrorString(rawValue))
                rawInput = rawValue
            default:
                // Number (type="n" or missing type attribute)
                if rawValue.isEmpty {
                    cellValue = .empty
                    rawInput = ""
                } else if let num = Double(rawValue) {
                    cellValue = .number(num)
                    rawInput = rawValue
                } else {
                    cellValue = .string(rawValue)
                    rawInput = rawValue
                }
            }
        }

        if !rawInput.isEmpty || cellValue != .empty {
            let cell = Cell(rawInput: rawInput, value: cellValue)
            cells[address] = cell
        }

        // Reset state
        currentCellRef = nil
        currentCellType = nil
    }

    private func resolveValue(rawValue: String, type: String) -> CellValue {
        if rawValue.isEmpty { return .empty }
        switch type {
        case "s":
            if let idx = Int(rawValue), idx >= 0, idx < sharedStrings.count {
                return .string(sharedStrings[idx])
            }
            return .string(rawValue)
        case "b":
            return .boolean(rawValue == "1")
        case "str":
            return .string(rawValue)
        case "e":
            return .error(mapErrorString(rawValue))
        default:
            if let num = Double(rawValue) {
                return .number(num)
            }
            return .string(rawValue)
        }
    }

    private func mapErrorString(_ errorStr: String) -> CellError {
        switch errorStr {
        case "#VALUE!": return .value
        case "#REF!": return .ref
        case "#DIV/0!": return .divZero
        case "#NAME?": return .name
        case "#N/A": return .na
        case "#NUM!": return .num
        default: return .generic
        }
    }
}
