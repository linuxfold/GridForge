import SwiftUI
import AppKit
import GridForgeCore

// MARK: - SheetTabsView

struct SheetTabsView: View {
    @ObservedObject var viewModel: WorkbookViewModel
    @State private var renamingIndex: Int? = nil
    @State private var renameText: String = ""
    @State private var hoveredIndex: Int? = nil

    var body: some View {
        HStack(spacing: 0) {
            // Add sheet button
            Button(action: {
                viewModel.addSheet()
            }) {
                Image(systemName: "plus")
                    .font(.system(size: 11, weight: .medium))
                    .foregroundColor(.secondary)
                    .frame(width: 28, height: GridForgeSpacing.sheetTabHeight)
                    .contentShape(Rectangle())
            }
            .buttonStyle(.plain)
            .padding(.leading, 4)

            // Tabs in scrollable area
            ScrollView(.horizontal, showsIndicators: false) {
                HStack(spacing: 2) {
                    ForEach(Array(viewModel.workbook.sheets.enumerated()), id: \.element.id) { index, sheet in
                        sheetTab(index: index, sheet: sheet)
                    }
                }
                .padding(.horizontal, 4)
            }

            Spacer()
        }
        .frame(height: GridForgeSpacing.sheetTabBarHeight)
        .background(GridForgeColors.toolbarBackground)
    }

    // MARK: - Individual Tab

    @ViewBuilder
    private func sheetTab(index: Int, sheet: Worksheet) -> some View {
        let isActive = index == viewModel.workbook.activeSheetIndex
        let isHovered = hoveredIndex == index

        Group {
            if renamingIndex == index {
                TextField("", text: $renameText)
                    .textFieldStyle(.plain)
                    .font(GridForgeTypography.sheetTabFont)
                    .frame(minWidth: GridForgeSpacing.sheetTabMinWidth, maxWidth: GridForgeSpacing.sheetTabMaxWidth)
                    .multilineTextAlignment(.center)
                    .onSubmit {
                        commitRename(at: index)
                    }
                    .onExitCommand {
                        renamingIndex = nil
                    }
                    .onAppear {
                        renameText = sheet.name
                    }
            } else {
                Text(sheet.name)
                    .font(GridForgeTypography.sheetTabFont)
                    .foregroundColor(isActive ? GridForgeColors.sheetTabActiveText : GridForgeColors.sheetTabInactiveText)
                    .lineLimit(1)
                    .truncationMode(.tail)
                    .frame(minWidth: GridForgeSpacing.sheetTabMinWidth, maxWidth: GridForgeSpacing.sheetTabMaxWidth)
            }
        }
        .padding(.horizontal, 12)
        .padding(.vertical, 4)
        .background(
            RoundedRectangle(cornerRadius: 6)
                .fill(tabBackground(isActive: isActive, isHovered: isHovered))
                .shadow(color: isActive ? Color.black.opacity(0.1) : Color.clear, radius: 1, y: 0.5)
        )
        .animation(GridForgeAnimation.quick, value: isHovered)
        .contentShape(Rectangle())
        .onHover { hovering in
            hoveredIndex = hovering ? index : nil
        }
        .onTapGesture(count: 2) {
            renameText = sheet.name
            renamingIndex = index
        }
        .onTapGesture(count: 1) {
            viewModel.switchSheet(to: index)
        }
        .contextMenu {
            Button("Rename") {
                renameText = sheet.name
                renamingIndex = index
            }
            Button("Duplicate") {
                viewModel.duplicateSheet(at: index)
            }

            Divider()

            Button("Move Left") {
                moveSheetLeft(at: index)
            }
            .disabled(index == 0)

            Button("Move Right") {
                moveSheetRight(at: index)
            }
            .disabled(index == viewModel.workbook.sheets.count - 1)

            if viewModel.workbook.sheets.count > 1 {
                Divider()
                Button("Delete", role: .destructive) {
                    viewModel.deleteSheet(at: index)
                }
            }
        }
    }

    // MARK: - Helpers

    private func tabBackground(isActive: Bool, isHovered: Bool) -> Color {
        if isActive {
            return GridForgeColors.sheetTabActive
        } else if isHovered {
            return GridForgeColors.sheetTabHover
        }
        return Color.clear
    }

    private func commitRename(at index: Int) {
        let trimmed = renameText.trimmingCharacters(in: .whitespaces)
        if !trimmed.isEmpty {
            viewModel.renameSheet(at: index, to: trimmed)
        }
        renamingIndex = nil
    }

    private func moveSheetLeft(at index: Int) {
        guard index > 0 else { return }
        viewModel.workbook.moveSheet(from: index, to: index - 1)
        if viewModel.workbook.activeSheetIndex == index {
            viewModel.workbook.activeSheetIndex = index - 1
        }
        viewModel.version += 1
    }

    private func moveSheetRight(at index: Int) {
        guard index < viewModel.workbook.sheets.count - 1 else { return }
        viewModel.workbook.moveSheet(from: index, to: index + 2)
        if viewModel.workbook.activeSheetIndex == index {
            viewModel.workbook.activeSheetIndex = index + 1
        }
        viewModel.version += 1
    }
}
