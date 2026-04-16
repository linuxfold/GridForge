import SwiftUI
import AppKit
import GridForgeCore
import UniformTypeIdentifiers

@main
struct GridForgeApp: App {
    var body: some Scene {
        WindowGroup {
            WorkbookWindowView()
                .frame(minWidth: 800, minHeight: 500)
        }
        .windowStyle(.titleBar)
        .windowResizability(.contentSize)
        .defaultSize(width: 1200, height: 800)
        .commands {
            GridForgeCommands()
        }
    }
}

// MARK: - App Commands

struct GridForgeCommands: Commands {
    @FocusedValue(\.activeWorkbookViewModel) var viewModel

    var body: some Commands {
        // File menu
        CommandGroup(replacing: .newItem) {
            Button("New Workbook") {
                viewModel?.newWorkbook()
            }
            .keyboardShortcut("n", modifiers: .command)

            Button("Open...") {
                guard let vm = viewModel else { return }
                let panel = NSOpenPanel()
                panel.allowedContentTypes = [UTType(filenameExtension: "xlsx")].compactMap { $0 }
                panel.canChooseDirectories = false
                if panel.runModal() == .OK, let url = panel.url {
                    vm.openFile(url: url)
                }
            }
            .keyboardShortcut("o", modifiers: .command)

            Divider()

            Button("Save") {
                guard let vm = viewModel else { return }
                if let url = vm.currentFileURL {
                    vm.saveFile(url: url)
                } else {
                    let panel = NSSavePanel()
                    panel.allowedContentTypes = [UTType(filenameExtension: "xlsx")].compactMap { $0 }
                    panel.nameFieldStringValue = "Untitled.xlsx"
                    if panel.runModal() == .OK, let url = panel.url {
                        vm.saveFile(url: url)
                    }
                }
            }
            .keyboardShortcut("s", modifiers: .command)

            Button("Save As...") {
                guard let vm = viewModel else { return }
                let panel = NSSavePanel()
                panel.allowedContentTypes = [
                    UTType(filenameExtension: "xlsx"),
                    UTType(filenameExtension: "csv")
                ].compactMap { $0 }
                let name = vm.currentFileURL?.lastPathComponent ?? "Untitled.xlsx"
                panel.nameFieldStringValue = name
                if panel.runModal() == .OK, let url = panel.url {
                    if url.pathExtension.lowercased() == "csv" {
                        vm.exportCSV(to: url)
                    } else {
                        vm.saveFile(url: url)
                    }
                }
            }
            .keyboardShortcut("s", modifiers: [.command, .shift])

            Divider()

            Button("Export as CSV...") {
                guard let vm = viewModel else { return }
                let panel = NSSavePanel()
                panel.allowedContentTypes = [UTType(filenameExtension: "csv")].compactMap { $0 }
                let baseName = vm.currentFileURL?.deletingPathExtension().lastPathComponent ?? "Untitled"
                panel.nameFieldStringValue = "\(baseName).csv"
                if panel.runModal() == .OK, let url = panel.url {
                    vm.exportCSV(to: url)
                }
            }
            .keyboardShortcut("e", modifiers: [.command, .shift])

            Button("Revert to Saved") {
                guard let vm = viewModel, vm.currentFileURL != nil else { return }
                let alert = NSAlert()
                alert.messageText = "Revert to Saved?"
                alert.informativeText = "Are you sure you want to revert? All unsaved changes will be lost."
                alert.addButton(withTitle: "Revert")
                alert.addButton(withTitle: "Cancel")
                if alert.runModal() == .alertFirstButtonReturn {
                    vm.revertToSaved()
                }
            }
            .disabled(viewModel?.currentFileURL == nil)
        }

        // Edit menu - Undo/Redo
        CommandGroup(replacing: .undoRedo) {
            Button("Undo") {
                viewModel?.undo()
            }
            .keyboardShortcut("z", modifiers: .command)
            .disabled(viewModel == nil || !(viewModel?.canUndo ?? false))

            Button("Redo") {
                viewModel?.redo()
            }
            .keyboardShortcut("z", modifiers: [.command, .shift])
            .disabled(viewModel == nil || !(viewModel?.canRedo ?? false))
        }

        // Edit menu - Pasteboard
        CommandGroup(replacing: .pasteboard) {
            Button("Cut") {
                viewModel?.cut()
            }
            .keyboardShortcut("x", modifiers: .command)
            .disabled(viewModel == nil)

            Button("Copy") {
                viewModel?.copy()
            }
            .keyboardShortcut("c", modifiers: .command)
            .disabled(viewModel == nil)

            Button("Paste") {
                viewModel?.paste()
            }
            .keyboardShortcut("v", modifiers: .command)
            .disabled(viewModel == nil)

            Divider()

            Button("Select All") {
                viewModel?.selectAll()
            }
            .keyboardShortcut("a", modifiers: .command)
            .disabled(viewModel == nil)

            Divider()

            Button("Find...") {
                // Placeholder for find bar toggle
            }
            .keyboardShortcut("f", modifiers: .command)
            .disabled(viewModel == nil)
        }

        // Format menu
        CommandMenu("Format") {
            Button("Bold") {
                viewModel?.toggleBold()
            }
            .keyboardShortcut("b", modifiers: .command)
            .disabled(viewModel == nil)

            Button("Italic") {
                viewModel?.toggleItalic()
            }
            .keyboardShortcut("i", modifiers: .command)
            .disabled(viewModel == nil)

            Button("Underline") {
                viewModel?.toggleUnderline()
            }
            .keyboardShortcut("u", modifiers: .command)
            .disabled(viewModel == nil)

            Divider()

            Button("Align Left") {
                viewModel?.setAlignment(.left)
            }
            .disabled(viewModel == nil)

            Button("Align Center") {
                viewModel?.setAlignment(.center)
            }
            .disabled(viewModel == nil)

            Button("Align Right") {
                viewModel?.setAlignment(.right)
            }
            .disabled(viewModel == nil)
        }

        // Sheet menu
        CommandMenu("Sheet") {
            Button("Add Sheet") {
                viewModel?.addSheet()
            }
            .disabled(viewModel == nil)

            Button("Duplicate Sheet") {
                guard let vm = viewModel else { return }
                vm.duplicateSheet(at: vm.workbook.activeSheetIndex)
            }
            .disabled(viewModel == nil)

            Button("Delete Sheet") {
                guard let vm = viewModel else { return }
                vm.deleteSheet(at: vm.workbook.activeSheetIndex)
            }
            .disabled(viewModel == nil || (viewModel?.workbook.sheets.count ?? 0) <= 1)

            Divider()

            Button("Insert Row") {
                viewModel?.insertRow()
            }
            .disabled(viewModel == nil)

            Button("Insert Column") {
                viewModel?.insertColumn()
            }
            .disabled(viewModel == nil)

            Divider()

            Button("Delete Row") {
                viewModel?.deleteRow()
            }
            .disabled(viewModel == nil)

            Button("Delete Column") {
                viewModel?.deleteColumn()
            }
            .disabled(viewModel == nil)
        }

        // View menu
        CommandGroup(after: .toolbar) {
            Divider()

            Button("Zoom In") {
                viewModel?.zoomIn()
            }
            .keyboardShortcut("=", modifiers: .command)
            .disabled(viewModel == nil)

            Button("Zoom Out") {
                viewModel?.zoomOut()
            }
            .keyboardShortcut("-", modifiers: .command)
            .disabled(viewModel == nil)

            Button("Actual Size") {
                viewModel?.zoomActualSize()
            }
            .keyboardShortcut("0", modifiers: .command)
            .disabled(viewModel == nil)

            Divider()

            Button(viewModel?.showFormulaBar == true ? "Hide Formula Bar" : "Show Formula Bar") {
                viewModel?.showFormulaBar.toggle()
            }
            .disabled(viewModel == nil)

            Button(viewModel?.showSheetTabs == true ? "Hide Sheet Tabs" : "Show Sheet Tabs") {
                viewModel?.showSheetTabs.toggle()
            }
            .disabled(viewModel == nil)

            Button(viewModel?.showStatusBar == true ? "Hide Status Bar" : "Show Status Bar") {
                viewModel?.showStatusBar.toggle()
            }
            .disabled(viewModel == nil)

            Divider()

            Button("Toggle Inspector") {
                viewModel?.showInspector.toggle()
            }
            .keyboardShortcut("0", modifiers: [.command, .option])
            .disabled(viewModel == nil)
        }

        // Help menu
        CommandGroup(replacing: .help) {
            Button("About GridForge") {
                let alert = NSAlert()
                alert.messageText = "GridForge"
                alert.informativeText = "A modern spreadsheet application for macOS.\nVersion 1.0"
                alert.alertStyle = .informational
                alert.addButton(withTitle: "OK")
                alert.runModal()
            }
        }
    }
}
