import SwiftUI
import AppKit
import GridForgeCore

// MARK: - FormulaBarView

struct FormulaBarView: View {
    @ObservedObject var viewModel: WorkbookViewModel
    @State private var isEditingNameBox: Bool = false
    @State private var nameBoxText: String = ""

    private var currentCell: Cell? {
        viewModel.activeSheet.cell(at: viewModel.activeCell)
    }

    var body: some View {
        VStack(spacing: 0) {
            HStack(spacing: 0) {
                // Name Box
                nameBox
                    .frame(minWidth: 70, maxWidth: 140)

                // 1px vertical divider
                Rectangle()
                    .fill(GridForgeColors.divider)
                    .frame(width: 1, height: 20)
                    .padding(.horizontal, 4)

                // Function icon
                functionIcon
                    .frame(width: 24)
                    .padding(.trailing, 4)

                // Formula text field
                formulaInput
            }
            .frame(height: GridForgeSpacing.formulaBarHeight)
            .padding(.horizontal, 8)
            .background(
                viewModel.isEditing
                    ? GridForgeColors.formulaBarBackground.opacity(0.95)
                    : GridForgeColors.formulaBarBackground
            )
            .overlay(
                // Subtle accent tint when editing
                viewModel.isEditing
                    ? Color.accentColor.opacity(0.04)
                    : Color.clear
            )

            // Bottom 1px separator
            Rectangle()
                .fill(GridForgeColors.formulaBarBorder)
                .frame(height: 1)
        }
    }

    // MARK: - Name Box

    @ViewBuilder
    private var nameBox: some View {
        if isEditingNameBox {
            TextField("", text: $nameBoxText)
                .textFieldStyle(.plain)
                .font(GridForgeTypography.nameBoxFont)
                .multilineTextAlignment(.center)
                .padding(.horizontal, 6)
                .padding(.vertical, 2)
                .background(
                    RoundedRectangle(cornerRadius: 3)
                        .fill(GridForgeColors.cellBackground)
                )
                .overlay(
                    RoundedRectangle(cornerRadius: 3)
                        .stroke(Color.accentColor, lineWidth: 1)
                )
                .onSubmit {
                    commitNameBox()
                }
                .onExitCommand {
                    isEditingNameBox = false
                }
                .onAppear {
                    nameBoxText = viewModel.activeCell.displayString
                }
        } else {
            Text(viewModel.activeCell.displayString)
                .font(GridForgeTypography.nameBoxFont)
                .frame(maxWidth: .infinity)
                .padding(.horizontal, 6)
                .padding(.vertical, 2)
                .background(
                    RoundedRectangle(cornerRadius: 3)
                        .fill(GridForgeColors.headerBackground.opacity(0.5))
                )
                .contentShape(Rectangle())
                .onTapGesture {
                    nameBoxText = viewModel.activeCell.displayString
                    isEditingNameBox = true
                }
        }
    }

    // MARK: - Function Icon

    @ViewBuilder
    private var functionIcon: some View {
        Image(systemName: "function")
            .font(.system(size: 12, weight: .medium))
            .foregroundColor(GridForgeColors.disabledText)
    }

    // MARK: - Formula Input

    @ViewBuilder
    private var formulaInput: some View {
        let displayText: String = {
            if viewModel.isEditing {
                return viewModel.editingText
            }
            return currentCell?.editString ?? ""
        }()

        TextField("", text: Binding(
            get: { displayText },
            set: { newValue in
                if !viewModel.isEditing {
                    viewModel.startEditing(withText: newValue)
                } else {
                    viewModel.editingText = newValue
                }
            }
        ))
        .textFieldStyle(.plain)
        .font(GridForgeTypography.formulaBarFont)
        .onSubmit {
            if viewModel.isEditing {
                viewModel.commitEdit()
            }
        }
        .onExitCommand {
            if viewModel.isEditing {
                viewModel.cancelEdit()
            }
        }
    }

    // MARK: - Helpers

    private func commitNameBox() {
        isEditingNameBox = false
        let trimmed = nameBoxText.trimmingCharacters(in: .whitespaces)
        if let addr = CellAddress.parse(trimmed) {
            viewModel.selectCell(addr)
        }
    }
}
