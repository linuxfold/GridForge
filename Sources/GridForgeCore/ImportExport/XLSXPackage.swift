import Foundation

/// Raw XLSX package data captured on import so unsupported content can survive save.
public struct XLSXPackage {
    public static let compatibilityPromise = "Open existing Excel files safely and preserve unsupported workbook content."

    public var entries: [String: Data]
    public var sheetPartPaths: [String]

    public init(entries: [String: Data], sheetPartPaths: [String]) {
        self.entries = entries
        self.sheetPartPaths = sheetPartPaths
    }

    public var containsUnsupportedContent: Bool {
        entries.keys.contains { path in
            !path.hasPrefix("xl/worksheets/")
                && path != "[Content_Types].xml"
                && path != "_rels/.rels"
                && path != "xl/workbook.xml"
                && path != "xl/_rels/workbook.xml.rels"
                && path != "xl/sharedStrings.xml"
                && path != "xl/styles.xml"
        }
    }
}
