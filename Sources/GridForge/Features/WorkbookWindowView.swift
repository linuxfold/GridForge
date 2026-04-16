import SwiftUI
import AppKit
import GridForgeCore
import UniformTypeIdentifiers

// MARK: - WorkbookWindowView

struct WorkbookWindowView: View {
    @StateObject var viewModel = WorkbookViewModel()

    var body: some View {
        VStack(spacing: 0) {
            // Formula bar
            if viewModel.showFormulaBar {
                FormulaBarView(viewModel: viewModel)
                Divider()
            }

            // Error banner
            if let error = viewModel.lastError {
                HStack(spacing: 6) {
                    Image(systemName: "exclamationmark.triangle.fill")
                        .foregroundColor(.white)
                        .font(.system(size: 11))
                    Text(error)
                        .font(GridForgeTypography.statusBarFont)
                        .foregroundColor(.white)
                        .lineLimit(1)
                    Spacer()
                    Button(action: { viewModel.lastError = nil }) {
                        Image(systemName: "xmark")
                            .font(.system(size: 9, weight: .bold))
                            .foregroundColor(.white.opacity(0.8))
                    }
                    .buttonStyle(.plain)
                }
                .padding(.horizontal, 8)
                .padding(.vertical, 4)
                .background(Color.red.opacity(0.85))
            }

            // Main content: grid + optional inspector
            ZStack {
                HStack(spacing: 0) {
                    SpreadsheetGridView(viewModel: viewModel)
                    if viewModel.showInspector {
                        GridForgeColors.divider
                            .frame(width: 1)
                        InspectorView(viewModel: viewModel)
                            .frame(width: GridForgeSpacing.inspectorWidth)
                    }
                }

                // Loading overlay
                if viewModel.isLoading {
                    Color.black.opacity(0.05)
                        .overlay(
                            ProgressView()
                                .scaleEffect(0.8)
                                .frame(width: 40, height: 40)
                                .background(
                                    RoundedRectangle(cornerRadius: 8)
                                        .fill(.ultraThinMaterial)
                                )
                        )
                        .allowsHitTesting(false)
                }
            }

            Divider()

            // Sheet tabs
            if viewModel.showSheetTabs {
                SheetTabsView(viewModel: viewModel)
            }

            // Status bar
            if viewModel.showStatusBar {
                StatusBarView(viewModel: viewModel)
            }
        }
        .navigationTitle(viewModel.windowTitle)
        .focusedSceneValue(\.activeWorkbookViewModel, viewModel)
        .toolbar {
            // File group
            ToolbarItemGroup(placement: .automatic) {
                Button(action: { viewModel.newWorkbook() }) {
                    Label("New", systemImage: "doc.badge.plus")
                }
                .help("New Workbook (Cmd+N)")

                Button(action: { openFile() }) {
                    Label("Open", systemImage: "folder")
                }
                .help("Open File (Cmd+O)")

                Button(action: { saveFile() }) {
                    Label("Save", systemImage: "square.and.arrow.down")
                }
                .help("Save File (Cmd+S)")

                Divider()

                // Format group
                Button(action: { viewModel.toggleBold() }) {
                    Label("Bold", systemImage: "bold")
                }
                .help("Toggle Bold (Cmd+B)")

                Button(action: { viewModel.toggleItalic() }) {
                    Label("Italic", systemImage: "italic")
                }
                .help("Toggle Italic (Cmd+I)")

                Button(action: { viewModel.toggleUnderline() }) {
                    Label("Underline", systemImage: "underline")
                }
                .help("Toggle Underline (Cmd+U)")

                Divider()

                // Edit group
                Button(action: { viewModel.undo() }) {
                    Label("Undo", systemImage: "arrow.uturn.backward")
                }
                .help("Undo (Cmd+Z)")
                .disabled(!viewModel.canUndo)

                Button(action: { viewModel.redo() }) {
                    Label("Redo", systemImage: "arrow.uturn.forward")
                }
                .help("Redo (Cmd+Shift+Z)")
                .disabled(!viewModel.canRedo)

                Divider()

                // View group
                Button(action: {
                    // Find placeholder: currently no-op, wired through menu
                }) {
                    Label("Find", systemImage: "magnifyingglass")
                }
                .help("Find (Cmd+F)")

                Button(action: { viewModel.showInspector.toggle() }) {
                    Label("Inspector", systemImage: "sidebar.right")
                }
                .help("Toggle Inspector (Cmd+Option+0)")
            }
        }
    }

    // MARK: File operations

    func openFile() {
        // Check for unsaved changes
        if viewModel.isDirty {
            let alert = NSAlert()
            alert.messageText = "Unsaved Changes"
            alert.informativeText = "You have unsaved changes. Do you want to save before opening a new file?"
            alert.addButton(withTitle: "Save")
            alert.addButton(withTitle: "Don't Save")
            alert.addButton(withTitle: "Cancel")
            let response = alert.runModal()
            switch response {
            case .alertFirstButtonReturn:
                saveFile()
            case .alertThirdButtonReturn:
                return
            default:
                break
            }
        }

        let panel = NSOpenPanel()
        panel.allowedContentTypes = [UTType(filenameExtension: "xlsx")].compactMap { $0 }
        panel.canChooseDirectories = false
        panel.allowsMultipleSelection = false
        if panel.runModal() == .OK, let url = panel.url {
            viewModel.openFile(url: url)
        }
    }

    func saveFile() {
        if let url = viewModel.currentFileURL {
            // Direct save to current file
            viewModel.saveFile(url: url)
        } else {
            saveFileAs()
        }
    }

    func saveFileAs() {
        let panel = NSSavePanel()
        panel.allowedContentTypes = [
            UTType(filenameExtension: "xlsx"),
            UTType(filenameExtension: "csv")
        ].compactMap { $0 }

        let fileName: String
        if let url = viewModel.currentFileURL {
            fileName = url.lastPathComponent
        } else {
            fileName = "Untitled.xlsx"
        }
        panel.nameFieldStringValue = fileName

        if panel.runModal() == .OK, let url = panel.url {
            if url.pathExtension.lowercased() == "csv" {
                viewModel.exportCSV(to: url)
            } else {
                viewModel.saveFile(url: url)
            }
        }
    }

    func exportCSV() {
        let panel = NSSavePanel()
        panel.allowedContentTypes = [UTType(filenameExtension: "csv")].compactMap { $0 }

        let baseName: String
        if let url = viewModel.currentFileURL {
            baseName = url.deletingPathExtension().lastPathComponent
        } else {
            baseName = "Untitled"
        }
        panel.nameFieldStringValue = "\(baseName).csv"

        if panel.runModal() == .OK, let url = panel.url {
            viewModel.exportCSV(to: url)
        }
    }
}

// MARK: - Status Bar

struct StatusBarView: View {
    @ObservedObject var viewModel: WorkbookViewModel

    var body: some View {
        HStack(spacing: 0) {
            Text(viewModel.statusText)
                .font(GridForgeTypography.statusBarFont)
                .foregroundColor(GridForgeColors.statusBarText)
                .lineLimit(1)
                .padding(.leading, 8)

            Spacer()

            if !viewModel.cellCountText.isEmpty {
                Text(viewModel.cellCountText)
                    .font(GridForgeTypography.statusBarFont)
                    .foregroundColor(GridForgeColors.statusBarText)
                    .lineLimit(1)
                    .padding(.trailing, 8)
            }

            if !viewModel.selectionSummary.isEmpty {
                GridForgeColors.divider
                    .frame(width: 1, height: 14)
                    .padding(.horizontal, 4)

                Text(viewModel.selectionSummary)
                    .font(GridForgeTypography.statusBarFont)
                    .foregroundColor(GridForgeColors.statusBarText)
                    .lineLimit(1)
                    .padding(.trailing, 8)
            }
        }
        .frame(height: GridForgeSpacing.statusBarHeight)
        .background(GridForgeColors.statusBarBackground)
    }
}

// MARK: - FocusedValue for ViewModel

struct ActiveWorkbookViewModelKey: FocusedValueKey {
    typealias Value = WorkbookViewModel
}

extension FocusedValues {
    var activeWorkbookViewModel: WorkbookViewModel? {
        get { self[ActiveWorkbookViewModelKey.self] }
        set { self[ActiveWorkbookViewModelKey.self] = newValue }
    }
}
