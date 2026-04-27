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

    private var formulaTextBinding: Binding<String> {
        Binding(
            get: {
                if viewModel.isEditing {
                    return viewModel.editingText
                }
                return currentCell?.editString ?? ""
            },
            set: { newValue in
                if !viewModel.isEditing {
                    viewModel.startEditing(withText: newValue)
                } else {
                    viewModel.editingText = newValue
                }
            }
        )
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
        FormulaEditorField(text: formulaTextBinding, viewModel: viewModel)
            .frame(maxWidth: .infinity, maxHeight: .infinity)
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

private struct FormulaEditorField: NSViewRepresentable {
    @Binding var text: String
    @ObservedObject var viewModel: WorkbookViewModel

    func makeCoordinator() -> Coordinator {
        Coordinator(parent: self)
    }

    func makeNSView(context: Context) -> NSTextField {
        let textField = NSTextField(frame: .zero)
        textField.delegate = context.coordinator
        textField.isBordered = false
        textField.isBezeled = false
        textField.drawsBackground = false
        textField.focusRingType = .none
        textField.font = GridForgeNSFonts.formulaBarFont
        textField.lineBreakMode = .byClipping
        textField.cell?.wraps = false
        textField.cell?.isScrollable = true
        textField.isEditable = true
        textField.isSelectable = true
        textField.stringValue = text
        return textField
    }

    func updateNSView(_ textField: NSTextField, context: Context) {
        context.coordinator.parent = self

        if let editor = textField.currentEditor() as? NSTextView {
            guard editor.string != text, !context.coordinator.isUpdatingFromField else { return }
            let selectedRange = editor.selectedRange()
            editor.string = text
            editor.setSelectedRange(NSRange(location: min(selectedRange.location, (text as NSString).length), length: 0))
        } else if textField.stringValue != text {
            textField.stringValue = text
        }
    }

    final class Coordinator: NSObject, NSTextFieldDelegate {
        var parent: FormulaEditorField
        var isUpdatingFromField = false

        init(parent: FormulaEditorField) {
            self.parent = parent
        }

        func controlTextDidBeginEditing(_ obj: Notification) {
            parent.viewModel.formulaBarHasFocus = true
            if !parent.viewModel.isEditing {
                parent.viewModel.startEditing(withText: parent.text)
            }
        }

        func controlTextDidChange(_ obj: Notification) {
            let value = fieldEditorText(from: obj) ?? (obj.object as? NSTextField)?.stringValue ?? parent.text
            isUpdatingFromField = true
            parent.text = value
            parent.viewModel.editingText = value
            isUpdatingFromField = false
        }

        func controlTextDidEndEditing(_ obj: Notification) {
            parent.viewModel.formulaBarHasFocus = false
        }

        func control(_ control: NSControl, textView: NSTextView, doCommandBy selector: Selector) -> Bool {
            if selector == #selector(NSResponder.insertNewline(_:)) {
                parent.text = textView.string
                parent.viewModel.editingText = textView.string
                parent.viewModel.commitEdit()
                parent.viewModel.formulaBarHasFocus = false
                NotificationCenter.default.post(name: .gridForgeFocusGrid, object: nil)
                return true
            }
            if selector == #selector(NSResponder.cancelOperation(_:)) {
                parent.viewModel.cancelEdit()
                parent.viewModel.formulaBarHasFocus = false
                NotificationCenter.default.post(name: .gridForgeFocusGrid, object: nil)
                return true
            }
            return false
        }

        private func fieldEditorText(from notification: Notification) -> String? {
            (notification.userInfo?["NSFieldEditor"] as? NSTextView)?.string
        }
    }
}
