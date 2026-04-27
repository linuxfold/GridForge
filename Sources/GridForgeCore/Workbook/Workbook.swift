import Foundation

/// Top-level document model: a collection of worksheets
public final class Workbook: Identifiable {
    public let id: UUID
    public var sheets: [Worksheet]
    public var activeSheetIndex: Int
    public var metadata: WorkbookMetadata
    public var sourceXLSXPackage: XLSXPackage?

    public init(
        sheets: [Worksheet]? = nil,
        metadata: WorkbookMetadata = WorkbookMetadata(),
        sourceXLSXPackage: XLSXPackage? = nil
    ) {
        self.id = UUID()
        self.sheets = sheets ?? [Worksheet(name: "Sheet1")]
        self.activeSheetIndex = 0
        self.metadata = metadata
        self.sourceXLSXPackage = sourceXLSXPackage
    }

    public var activeSheet: Worksheet {
        sheets[activeSheetIndex]
    }

    // MARK: - Sheet Management

    @discardableResult
    public func addSheet(name: String? = nil) -> Worksheet {
        let sheetName = name ?? nextSheetName()
        let sheet = Worksheet(name: sheetName)
        sheets.append(sheet)
        return sheet
    }

    public func deleteSheet(at index: Int) {
        guard sheets.count > 1, index < sheets.count else { return }
        sheets.remove(at: index)
        if activeSheetIndex >= sheets.count {
            activeSheetIndex = sheets.count - 1
        }
    }

    public func renameSheet(at index: Int, to name: String) {
        guard index < sheets.count else { return }
        sheets[index].name = name
    }

    @discardableResult
    public func duplicateSheet(at index: Int) -> Worksheet? {
        guard index < sheets.count else { return nil }
        let source = sheets[index]
        let copy = Worksheet(name: "\(source.name) Copy")
        for (addr, cell) in source.cells {
            copy.cells[addr] = cell.copy()
        }
        copy.columnWidths = source.columnWidths
        copy.rowHeights = source.rowHeights
        sheets.insert(copy, at: index + 1)
        return copy
    }

    public func moveSheet(from: Int, to: Int) {
        guard from < sheets.count, to <= sheets.count, from != to else { return }
        let sheet = sheets.remove(at: from)
        let dest = to > from ? to - 1 : to
        sheets.insert(sheet, at: dest)
        if activeSheetIndex == from {
            activeSheetIndex = dest
        }
    }

    private func nextSheetName() -> String {
        var idx = sheets.count + 1
        while sheets.contains(where: { $0.name == "Sheet\(idx)" }) {
            idx += 1
        }
        return "Sheet\(idx)"
    }
}

public struct WorkbookMetadata: Equatable, Sendable {
    public var title: String
    public var author: String
    public var createdDate: Date
    public var modifiedDate: Date

    public init(
        title: String = "Untitled",
        author: String = "",
        createdDate: Date = Date(),
        modifiedDate: Date = Date()
    ) {
        self.title = title
        self.author = author
        self.createdDate = createdDate
        self.modifiedDate = modifiedDate
    }
}
