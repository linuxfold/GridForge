import SwiftUI
import UniformTypeIdentifiers
import GridForgeCore

struct GridForgeDocument: FileDocument {
    static var readableContentTypes: [UTType] { [.gridForgeXLSX] }
    static var writableContentTypes: [UTType] { [.gridForgeXLSX] }

    var workbook: Workbook

    init(workbook: Workbook = Workbook()) {
        self.workbook = workbook
    }

    init(configuration: ReadConfiguration) throws {
        guard let data = configuration.file.regularFileContents else {
            throw CocoaError(.fileReadCorruptFile)
        }
        let url = Self.temporaryXLSXURL()
        try data.write(to: url, options: .atomic)
        defer { try? FileManager.default.removeItem(at: url) }
        self.workbook = try XLSXReader.read(from: url)
    }

    func fileWrapper(configuration: WriteConfiguration) throws -> FileWrapper {
        let url = Self.temporaryXLSXURL()
        defer { try? FileManager.default.removeItem(at: url) }
        try XLSXWriter.write(workbook, to: url)
        let data = try Data(contentsOf: url)
        return FileWrapper(regularFileWithContents: data)
    }

    private static func temporaryXLSXURL() -> URL {
        FileManager.default.temporaryDirectory
            .appendingPathComponent("GridForge-\(UUID().uuidString)")
            .appendingPathExtension("xlsx")
    }
}

extension UTType {
    static let gridForgeXLSX = UTType(
        "org.openxmlformats.spreadsheetml.sheet"
    ) ?? UTType(
        filenameExtension: "xlsx",
        conformingTo: .zip
    )!
}
