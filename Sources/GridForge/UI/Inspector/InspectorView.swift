import SwiftUI
import AppKit
import GridForgeCore

// MARK: - InspectorView

struct InspectorView: View {
    @ObservedObject var viewModel: WorkbookViewModel

    private var activeCell: CellAddress { viewModel.activeCell }
    private var sheet: Worksheet { viewModel.activeSheet }
    private var cell: Cell? { sheet.cell(at: activeCell) }

    var body: some View {
        ScrollView {
            VStack(alignment: .leading, spacing: 12) {
                cellInfoSection
                fontAndTextSection
                colorsSection
                sheetInfoSection
            }
            .padding(12)
        }
        .frame(maxHeight: .infinity)
        .background(GridForgeColors.inspectorBackground)
    }

    // MARK: - Section 1: Cell Info

    @ViewBuilder
    private var cellInfoSection: some View {
        DisclosureGroup {
            VStack(alignment: .leading, spacing: 8) {
                // Cell address
                HStack {
                    Text("Address")
                        .font(GridForgeTypography.inspectorLabel)
                        .foregroundColor(.secondary)
                        .frame(width: 70, alignment: .leading)
                    Text(activeCell.displayString)
                        .font(GridForgeTypography.inspectorValue)
                        .bold()
                        .textSelection(.enabled)
                }

                // Value type badge
                HStack {
                    Text("Type")
                        .font(GridForgeTypography.inspectorLabel)
                        .foregroundColor(.secondary)
                        .frame(width: 70, alignment: .leading)
                    valueTypeBadge
                }

                // Raw input
                if let cell = cell, !cell.rawInput.isEmpty {
                    HStack(alignment: .top) {
                        Text("Input")
                            .font(GridForgeTypography.inspectorLabel)
                            .foregroundColor(.secondary)
                            .frame(width: 70, alignment: .leading)
                        Text(cell.rawInput)
                            .font(cell.isFormula ? GridForgeTypography.inspectorValue : GridForgeTypography.inspectorLabel)
                            .foregroundColor(cell.isFormula ? GridForgeColors.formulaText : GridForgeColors.cellText)
                            .textSelection(.enabled)
                            .frame(maxWidth: .infinity, alignment: .leading)
                    }
                }

                // Computed value
                if let cell = cell, !cell.isEmpty {
                    HStack(alignment: .top) {
                        Text("Value")
                            .font(GridForgeTypography.inspectorLabel)
                            .foregroundColor(.secondary)
                            .frame(width: 70, alignment: .leading)
                        Text(cell.displayString)
                            .font(GridForgeTypography.inspectorValue)
                            .textSelection(.enabled)
                            .frame(maxWidth: .infinity, alignment: .leading)
                    }
                }

                // Error explanation
                if let cell = cell, case .error(let err) = cell.value {
                    HStack(alignment: .top) {
                        Text("Error")
                            .font(GridForgeTypography.inspectorLabel)
                            .foregroundColor(.secondary)
                            .frame(width: 70, alignment: .leading)
                        Text(errorExplanation(err))
                            .font(GridForgeTypography.inspectorLabel)
                            .foregroundColor(GridForgeColors.errorText)
                            .frame(maxWidth: .infinity, alignment: .leading)
                    }
                }
            }
            .frame(maxWidth: .infinity, alignment: .leading)
            .padding(.top, 4)
        } label: {
            Label("Cell Info", systemImage: "tablecells")
                .font(.headline)
        }
    }

    // MARK: - Section 2: Font & Text

    @ViewBuilder
    private var fontAndTextSection: some View {
        DisclosureGroup {
            VStack(alignment: .leading, spacing: 10) {
                // Bold / Italic / Underline toggles
                HStack(spacing: 6) {
                    Toggle(isOn: Binding(
                        get: { cell?.formatting.bold ?? false },
                        set: { _ in viewModel.toggleBold() }
                    )) {
                        Image(systemName: "bold")
                    }
                    .toggleStyle(.button)

                    Toggle(isOn: Binding(
                        get: { cell?.formatting.italic ?? false },
                        set: { _ in viewModel.toggleItalic() }
                    )) {
                        Image(systemName: "italic")
                    }
                    .toggleStyle(.button)

                    Toggle(isOn: Binding(
                        get: { cell?.formatting.underline ?? false },
                        set: { _ in viewModel.toggleUnderline() }
                    )) {
                        Image(systemName: "underline")
                    }
                    .toggleStyle(.button)

                    Spacer()
                }

                Divider()

                // Font size stepper
                HStack {
                    Text("Font Size")
                        .font(GridForgeTypography.inspectorLabel)
                        .foregroundColor(.secondary)
                    Spacer()
                    Text("\(Int(cell?.formatting.fontSize ?? 13)) pt")
                        .font(GridForgeTypography.inspectorValue)
                        .monospacedDigit()
                    Stepper("", value: Binding(
                        get: { cell?.formatting.fontSize ?? 13 },
                        set: { viewModel.setFontSize($0) }
                    ), in: 11...72, step: 1)
                    .labelsHidden()
                }

                Divider()

                // Alignment segmented control
                Text("Alignment")
                    .font(GridForgeTypography.inspectorLabel)
                    .foregroundColor(.secondary)

                Picker("", selection: Binding(
                    get: { cell?.formatting.alignment ?? .general },
                    set: { viewModel.setAlignment($0) }
                )) {
                    Image(systemName: "textformat.justifyleft")
                        .tag(GridForgeCore.HorizontalAlignment.general)
                    Image(systemName: "text.alignleft")
                        .tag(GridForgeCore.HorizontalAlignment.left)
                    Image(systemName: "text.aligncenter")
                        .tag(GridForgeCore.HorizontalAlignment.center)
                    Image(systemName: "text.alignright")
                        .tag(GridForgeCore.HorizontalAlignment.right)
                }
                .pickerStyle(.segmented)
            }
            .frame(maxWidth: .infinity, alignment: .leading)
            .padding(.top, 4)
        } label: {
            Label("Font & Text", systemImage: "textformat")
                .font(.headline)
        }
    }

    // MARK: - Section 3: Colors

    @ViewBuilder
    private var colorsSection: some View {
        DisclosureGroup {
            VStack(alignment: .leading, spacing: 8) {
                // Text color
                HStack(spacing: 8) {
                    Circle()
                        .fill(textColorDisplay)
                        .frame(width: 14, height: 14)
                        .overlay(
                            Circle()
                                .stroke(Color.secondary.opacity(0.3), lineWidth: 0.5)
                        )
                    Text("Text Color")
                        .font(GridForgeTypography.inspectorLabel)
                    Spacer()
                    Text(cell?.formatting.textColor != nil ? "Custom" : "Automatic")
                        .font(GridForgeTypography.inspectorLabel)
                        .foregroundColor(.secondary)
                }

                // Background color
                HStack(spacing: 8) {
                    Circle()
                        .fill(backgroundColorDisplay)
                        .frame(width: 14, height: 14)
                        .overlay(
                            Circle()
                                .stroke(Color.secondary.opacity(0.3), lineWidth: 0.5)
                        )
                    Text("Background")
                        .font(GridForgeTypography.inspectorLabel)
                    Spacer()
                    Text(cell?.formatting.backgroundColor != nil ? "Custom" : "Automatic")
                        .font(GridForgeTypography.inspectorLabel)
                        .foregroundColor(.secondary)
                }
            }
            .frame(maxWidth: .infinity, alignment: .leading)
            .padding(.top, 4)
        } label: {
            Label("Colors", systemImage: "paintpalette")
                .font(.headline)
        }
    }

    // MARK: - Section 4: Sheet Info

    @ViewBuilder
    private var sheetInfoSection: some View {
        DisclosureGroup {
            VStack(alignment: .leading, spacing: 8) {
                // Sheet name (editable)
                HStack {
                    Text("Name")
                        .font(GridForgeTypography.inspectorLabel)
                        .foregroundColor(.secondary)
                        .frame(width: 70, alignment: .leading)
                    TextField("Sheet name", text: Binding(
                        get: { sheet.name },
                        set: { newName in
                            let trimmed = newName.trimmingCharacters(in: .whitespaces)
                            if !trimmed.isEmpty {
                                viewModel.renameSheet(at: viewModel.workbook.activeSheetIndex, to: trimmed)
                            }
                        }
                    ))
                    .textFieldStyle(.plain)
                    .font(GridForgeTypography.inspectorValue)
                }

                // Cells used
                HStack {
                    Text("Cells Used")
                        .font(GridForgeTypography.inspectorLabel)
                        .foregroundColor(.secondary)
                        .frame(width: 70, alignment: .leading)
                    Text("\(sheet.cells.count)")
                        .font(GridForgeTypography.inspectorValue)
                        .monospacedDigit()
                }

                // Used range
                if let usedRange = sheet.usedRange {
                    HStack {
                        Text("Range")
                            .font(GridForgeTypography.inspectorLabel)
                            .foregroundColor(.secondary)
                            .frame(width: 70, alignment: .leading)
                        Text(usedRange.displayString)
                            .font(GridForgeTypography.inspectorValue)
                            .textSelection(.enabled)
                    }
                }
            }
            .frame(maxWidth: .infinity, alignment: .leading)
            .padding(.top, 4)
        } label: {
            Label("Sheet Info", systemImage: "doc.text")
                .font(.headline)
        }
    }

    // MARK: - Helpers

    @ViewBuilder
    private var valueTypeBadge: some View {
        let (label, color) = cellTypeInfo
        Text(label)
            .font(.system(size: 10, weight: .medium))
            .foregroundColor(.white)
            .padding(.horizontal, 6)
            .padding(.vertical, 2)
            .background(
                RoundedRectangle(cornerRadius: 4)
                    .fill(color)
            )
    }

    private var cellTypeInfo: (String, Color) {
        guard let cell = cell else { return ("Empty", Color.gray) }
        switch cell.value {
        case .empty: return ("Empty", Color.gray)
        case .string: return ("Text", Color.blue)
        case .number: return ("Number", Color.green)
        case .boolean: return ("Boolean", Color.orange)
        case .date: return ("Date", Color.purple)
        case .error: return ("Error", Color.red)
        }
    }

    private var textColorDisplay: Color {
        if let cc = cell?.formatting.textColor {
            return Color(red: cc.red, green: cc.green, blue: cc.blue, opacity: cc.alpha)
        }
        return GridForgeColors.cellText
    }

    private var backgroundColorDisplay: Color {
        if let cc = cell?.formatting.backgroundColor {
            return Color(red: cc.red, green: cc.green, blue: cc.blue, opacity: cc.alpha)
        }
        return GridForgeColors.cellBackground
    }

    private func errorExplanation(_ error: CellError) -> String {
        switch error {
        case .value: return "A value in the formula is the wrong type."
        case .ref: return "A cell reference is invalid."
        case .divZero: return "Division by zero."
        case .name: return "Unrecognized formula name."
        case .na: return "Value not available."
        case .circular: return "Circular reference detected."
        case .generic: return "An error occurred."
        case .syntax: return "Formula syntax error."
        case .num: return "Invalid numeric value."
        }
    }
}
