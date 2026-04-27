import SwiftUI
import AppKit
import GridForgeCore

extension Notification.Name {
    static let gridForgeFocusGrid = Notification.Name("GridForgeFocusGrid")
}

// MARK: - SpreadsheetGridView

struct SpreadsheetGridView: NSViewRepresentable {
    @ObservedObject var viewModel: WorkbookViewModel

    func makeCoordinator() -> Coordinator {
        Coordinator(viewModel: viewModel)
    }

    func makeNSView(context: Context) -> GridContainerView {
        GridContainerView(coordinator: context.coordinator)
    }

    func updateNSView(_ container: GridContainerView, context: Context) {
        context.coordinator.viewModel = viewModel
        if viewModel.formulaBarHasFocus {
            context.coordinator.discardFormulaReferenceInsertionState()
            context.coordinator.removeEditor()
        } else if viewModel.isEditing {
            context.coordinator.syncEditorFromViewModelIfNeeded()
        } else {
            context.coordinator.removeEditor()
        }
        container.updateDocumentSize()
        container.invalidateGrid()
    }
}

fileprivate struct FormulaReferenceHighlight {
    let cellRange: CellRange
    let color: NSColor
    let textRange: NSRange
}

// MARK: - Coordinator

final class Coordinator: NSObject, NSTextFieldDelegate {
    var viewModel: WorkbookViewModel
    weak var activeEditor: NSTextField?
    weak var scrollView: NSScrollView?
    weak var containerView: GridContainerView?
    private var isApplyingEditorAttributes = false

    var hoverCell: CellAddress?
    var hoverColumn: Int?
    var hoverRow: Int?
    var isResizingColumn = false
    var resizeColumnIndex = -1
    var resizeStartX: CGFloat = 0
    var resizeStartWidth: CGFloat = 0
    var resizePreviewWidth: CGFloat = 0
    var selectionDragAnchor: CellAddress?
    var formulaRangeAnchor: CellAddress?
    var formulaInsertedTextRange: NSRange?
    var formulaReplaceableTextRange: NSRange?

    let cellAttrsNormal: [NSAttributedString.Key: Any]
    let cellAttrsBold: [NSAttributedString.Key: Any]
    let cellAttrsItalic: [NSAttributedString.Key: Any]
    let cellAttrsBoldItalic: [NSAttributedString.Key: Any]
    let headerAttrs: [NSAttributedString.Key: Any]
    let activeHeaderAttrs: [NSAttributedString.Key: Any]
    let formulaReferenceColors: [NSColor] = [
        .systemBlue,
        .systemRed,
        .systemGreen,
        .systemPurple,
        .systemOrange,
        .systemTeal
    ]

    private static let formulaReferenceRegex = try? NSRegularExpression(
        pattern: "(?<![A-Za-z0-9_$])\\$?[A-Za-z]{1,3}\\$?[0-9]+(?::\\$?[A-Za-z]{1,3}\\$?[0-9]+)?(?![A-Za-z0-9_\\(])"
    )

    init(viewModel: WorkbookViewModel) {
        self.viewModel = viewModel

        let centerPara = NSMutableParagraphStyle()
        centerPara.alignment = .center
        centerPara.lineBreakMode = .byClipping
        headerAttrs = [
            .font: GridForgeNSFonts.headerFont,
            .foregroundColor: GridForgeNSColors.headerText,
            .paragraphStyle: centerPara
        ]
        activeHeaderAttrs = [
            .font: NSFont.systemFont(ofSize: 11, weight: .bold),
            .foregroundColor: NSColor.controlAccentColor,
            .paragraphStyle: centerPara
        ]

        let clipPara = NSMutableParagraphStyle()
        clipPara.lineBreakMode = .byClipping
        cellAttrsNormal = [
            .font: GridForgeNSFonts.cellFont,
            .foregroundColor: GridForgeNSColors.cellText,
            .paragraphStyle: clipPara
        ]
        cellAttrsBold = [
            .font: GridForgeNSFonts.cellFontBold,
            .foregroundColor: GridForgeNSColors.cellText,
            .paragraphStyle: clipPara
        ]
        cellAttrsItalic = [
            .font: GridForgeNSFonts.cellFontItalic,
            .foregroundColor: GridForgeNSColors.cellText,
            .paragraphStyle: clipPara
        ]
        cellAttrsBoldItalic = [
            .font: GridForgeNSFonts.cellFontBoldItalic,
            .foregroundColor: GridForgeNSColors.cellText,
            .paragraphStyle: clipPara
        ]
        super.init()
    }

    var rhw: CGFloat { GridForgeSpacing.rowHeaderWidth }
    var chh: CGFloat { GridForgeSpacing.columnHeaderHeight }
    var sheet: Worksheet { viewModel.activeSheet }
    var nCols: Int { viewModel.displayColumns }
    var nRows: Int { viewModel.displayRows }
    var isEditingFormula: Bool {
        viewModel.isEditing && viewModel.editingText.trimmingCharacters(in: .whitespacesAndNewlines).hasPrefix("=")
    }

    func xForCol(_ column: Int) -> CGFloat {
        CGFloat(sheet.xOffset(for: column, totalColumns: nCols))
    }

    func yForRow(_ row: Int) -> CGFloat {
        CGFloat(sheet.yOffset(for: row, totalRows: nRows))
    }

    func colW(_ column: Int) -> CGFloat {
        CGFloat(sheet.columnWidth(for: column))
    }

    func rowH(_ row: Int) -> CGFloat {
        CGFloat(sheet.rowHeight(for: row))
    }

    func colAt(_ x: CGFloat) -> Int {
        sheet.columnAt(x: Double(max(0, x)), totalColumns: nCols)
    }

    func rowAt(_ y: CGFloat) -> Int {
        sheet.rowAt(y: Double(max(0, y)), totalRows: nRows)
    }

    func visibleCols(in rect: NSRect) -> ClosedRange<Int> {
        colAt(rect.minX)...colAt(max(rect.minX, rect.maxX))
    }

    func visibleRows(in rect: NSRect) -> ClosedRange<Int> {
        rowAt(rect.minY)...rowAt(max(rect.minY, rect.maxY))
    }

    func colBorderHit(atContentX x: CGFloat, tol: CGFloat = 4) -> Int? {
        guard let sv = scrollView else { return nil }
        let visible = NSRect(x: sv.documentVisibleRect.minX, y: 0, width: sv.documentVisibleRect.width, height: chh)
        for column in visibleCols(in: visible) {
            if abs(x - (xForCol(column) + colW(column))) <= tol {
                return column
            }
        }
        return nil
    }

    func cellAttributes(for cell: Cell) -> [NSAttributedString.Key: Any] {
        var attrs = (cell.formatting.bold && cell.formatting.italic) ? cellAttrsBoldItalic
            : cell.formatting.bold ? cellAttrsBold
            : cell.formatting.italic ? cellAttrsItalic : cellAttrsNormal

        if cell.value.isError {
            attrs[.foregroundColor] = GridForgeNSColors.errorText
        }

        let paragraph = NSMutableParagraphStyle()
        paragraph.lineBreakMode = .byClipping
        switch cell.formatting.alignment {
        case .left:
            paragraph.alignment = .left
        case .center:
            paragraph.alignment = .center
        case .right:
            paragraph.alignment = .right
        case .general:
            paragraph.alignment = cell.value.numericValue == nil ? .left : .right
        }
        attrs[.paragraphStyle] = paragraph
        return attrs
    }

    fileprivate func formulaReferenceHighlights() -> [FormulaReferenceHighlight] {
        guard isEditingFormula, let regex = Self.formulaReferenceRegex else { return [] }

        let text = viewModel.editingText
        let nsText = text as NSString
        let matches = regex.matches(in: text, range: NSRange(location: 0, length: nsText.length))

        var highlights: [FormulaReferenceHighlight] = []
        highlights.reserveCapacity(matches.count)

        for match in matches {
            let token = nsText.substring(with: match.range)
            let parts = token.split(separator: ":", maxSplits: 1).map(String.init)
            guard let start = CellReference.parse(parts[0]) else { continue }
            let end = parts.count == 2 ? CellReference.parse(parts[1]) : start
            guard let end else { continue }

            let color = formulaReferenceColors[highlights.count % formulaReferenceColors.count]
            highlights.append(FormulaReferenceHighlight(
                cellRange: CellRange(start: start.address, end: end.address),
                color: color,
                textRange: match.range
            ))
        }

        return highlights
    }

    func insertFormulaReference(_ address: CellAddress) {
        guard isEditingFormula else { return }
        formulaRangeAnchor = address
        formulaInsertedTextRange = insertTextIntoFormulaEditor(
            address.displayString,
            replacing: formulaReplaceableTextRange,
            keepReferenceReplaceable: true
        )
        formulaReplaceableTextRange = formulaInsertedTextRange
    }

    func updateFormulaRangeReference(to address: CellAddress) {
        guard let anchor = formulaRangeAnchor, let insertedRange = formulaInsertedTextRange else { return }
        let rangeText = CellRange(start: anchor, end: address).displayString
        formulaInsertedTextRange = insertTextIntoFormulaEditor(
            rangeText,
            replacing: insertedRange,
            keepReferenceReplaceable: true
        )
        formulaReplaceableTextRange = formulaInsertedTextRange
    }

    func endFormulaRangeReference() {
        formulaRangeAnchor = nil
        formulaInsertedTextRange = nil
    }

    func discardFormulaReferenceInsertionState() {
        formulaRangeAnchor = nil
        formulaInsertedTextRange = nil
        formulaReplaceableTextRange = nil
    }

    @discardableResult
    func insertTextIntoFormulaEditor(
        _ text: String,
        replacing replacementRange: NSRange? = nil,
        keepReferenceReplaceable: Bool = false
    ) -> NSRange? {
        guard isEditingFormula, !text.isEmpty else { return nil }
        if !keepReferenceReplaceable {
            discardFormulaReferenceInsertionState()
        }

        if let editor = activeEditor {
            if let fieldEditor = editor.currentEditor() as? NSTextView {
                let currentText = fieldEditor.string as NSString
                let selectedRange = boundedRange(replacementRange ?? fieldEditor.selectedRange(), in: currentText.length)
                let replacement = currentText.replacingCharacters(in: selectedRange, with: text)

                viewModel.editingText = replacement
                editor.stringValue = replacement
                isApplyingEditorAttributes = true
                fieldEditor.string = replacement
                isApplyingEditorAttributes = false
                fieldEditor.setSelectedRange(NSRange(location: selectedRange.location + (text as NSString).length, length: 0))
                editor.window?.makeFirstResponder(fieldEditor)
                applyEditorTextAttributes()
                invalidateGrid()
                return NSRange(location: selectedRange.location, length: (text as NSString).length)
            } else {
                let currentText = viewModel.editingText as NSString
                let selectedRange = boundedRange(replacementRange ?? NSRange(location: currentText.length, length: 0), in: currentText.length)
                let replacement = currentText.replacingCharacters(in: selectedRange, with: text)

                viewModel.editingText = replacement
                editor.stringValue = replacement
                applyEditorTextAttributes()
                invalidateGrid()
                return NSRange(location: selectedRange.location, length: (text as NSString).length)
            }
        } else {
            let currentText = viewModel.editingText as NSString
            let selectedRange = boundedRange(replacementRange ?? NSRange(location: currentText.length, length: 0), in: currentText.length)
            viewModel.editingText = currentText.replacingCharacters(in: selectedRange, with: text)
            invalidateGrid()
            return NSRange(location: selectedRange.location, length: (text as NSString).length)
        }
    }

    func applyEditorTextAttributes() {
        guard let editor = activeEditor, !isApplyingEditorAttributes else { return }

        let text = viewModel.editingText
        let nsText = text as NSString
        let attributed = NSMutableAttributedString(
            string: text,
            attributes: [
                .font: GridForgeNSFonts.editorFont,
                .foregroundColor: GridForgeNSColors.cellText
            ]
        )

        if isEditingFormula {
            for highlight in formulaReferenceHighlights() {
                attributed.addAttribute(.foregroundColor, value: highlight.color, range: highlight.textRange)
            }
        }

        isApplyingEditorAttributes = true
        if let fieldEditor = editor.currentEditor() as? NSTextView {
            let selectedRange = boundedRange(fieldEditor.selectedRange(), in: nsText.length)
            fieldEditor.textStorage?.setAttributedString(attributed)
            fieldEditor.setSelectedRange(selectedRange)
        } else {
            editor.attributedStringValue = attributed
        }
        isApplyingEditorAttributes = false
    }

    private func boundedRange(_ range: NSRange, in length: Int) -> NSRange {
        let location = min(max(0, range.location), length)
        let maxLength = max(0, length - location)
        return NSRange(location: location, length: min(range.length, maxLength))
    }

    func beginEditing(in bodyView: GridBodyView, withText text: String? = nil) {
        removeEditor()
        let address = viewModel.activeCell
        viewModel.startEditing(withText: text)

        let editorRect = NSRect(
            x: xForCol(address.column) - 1,
            y: yForRow(address.row) - 1,
            width: colW(address.column) + 2,
            height: rowH(address.row) + 2
        )
        let editor = NSTextField(frame: editorRect)
        editor.font = GridForgeNSFonts.editorFont
        editor.isBordered = false
        editor.isBezeled = false
        editor.focusRingType = .none
        editor.drawsBackground = true
        editor.backgroundColor = GridForgeNSColors.editorBackground
        editor.wantsLayer = true
        editor.layer?.borderColor = GridForgeNSColors.editorBorder.cgColor
        editor.layer?.borderWidth = 2
        editor.layer?.cornerRadius = 1
        editor.layer?.shadowColor = GridForgeNSColors.editorShadow.cgColor
        editor.layer?.shadowOffset = CGSize(width: 0, height: -1)
        editor.layer?.shadowRadius = 3
        editor.layer?.shadowOpacity = 1
        editor.delegate = self
        editor.stringValue = viewModel.editingText
        editor.isEditable = true
        editor.isSelectable = true

        bodyView.addSubview(editor)
        editor.selectText(nil)
        if let text, !text.isEmpty, let fieldEditor = editor.currentEditor() {
            fieldEditor.selectedRange = NSRange(location: text.count, length: 0)
        }
        activeEditor = editor
        applyEditorTextAttributes()
    }

    func removeEditor() {
        activeEditor?.removeFromSuperview()
        activeEditor = nil
    }

    func syncEditorFromViewModelIfNeeded() {
        guard let editor = activeEditor, editor.currentEditor() == nil else { return }
        guard editor.stringValue != viewModel.editingText else { return }
        editor.stringValue = viewModel.editingText
        applyEditorTextAttributes()
    }

    func commitAndRemoveEditor() {
        if let editor = activeEditor {
            if let fieldEditor = editor.currentEditor() as? NSTextView {
                viewModel.editingText = fieldEditor.string
            } else {
                viewModel.editingText = editor.stringValue
            }
        }
        viewModel.commitEdit()
        removeEditor()
        focusGrid()
        invalidateGrid()
    }

    func cancelAndRemoveEditor() {
        viewModel.cancelEdit()
        removeEditor()
        focusGrid()
        invalidateGrid()
    }

    func scrollToActiveCell() {
        guard let scrollView else { return }
        let address = viewModel.activeCell
        let cellRect = NSRect(
            x: xForCol(address.column),
            y: yForRow(address.row),
            width: colW(address.column),
            height: rowH(address.row)
        )
        let visible = scrollView.documentVisibleRect
        var origin = visible.origin

        if cellRect.maxX > visible.maxX {
            origin.x = cellRect.maxX - visible.width
        }
        if cellRect.minX < visible.minX {
            origin.x = cellRect.minX
        }
        if cellRect.maxY > visible.maxY {
            origin.y = cellRect.maxY - visible.height
        }
        if cellRect.minY < visible.minY {
            origin.y = cellRect.minY
        }

        origin.x = max(0, origin.x)
        origin.y = max(0, origin.y)
        if origin != visible.origin {
            scrollView.contentView.scroll(to: origin)
            scrollView.reflectScrolledClipView(scrollView.contentView)
            containerView?.syncHeaderOffsets()
        }
    }

    func buildContextMenu() -> NSMenu {
        let menu = NSMenu()
        menu.addItem(withTitle: "Cut", action: #selector(GridBodyView.ctxCut(_:)), keyEquivalent: "x")
        menu.addItem(withTitle: "Copy", action: #selector(GridBodyView.ctxCopy(_:)), keyEquivalent: "c")
        menu.addItem(withTitle: "Paste", action: #selector(GridBodyView.ctxPaste(_:)), keyEquivalent: "v")
        menu.addItem(.separator())
        menu.addItem(withTitle: "Insert Row Above", action: #selector(GridBodyView.ctxInsRowAbove(_:)), keyEquivalent: "")
        menu.addItem(withTitle: "Insert Row Below", action: #selector(GridBodyView.ctxInsRowBelow(_:)), keyEquivalent: "")
        menu.addItem(withTitle: "Insert Column Left", action: #selector(GridBodyView.ctxInsColLeft(_:)), keyEquivalent: "")
        menu.addItem(withTitle: "Insert Column Right", action: #selector(GridBodyView.ctxInsColRight(_:)), keyEquivalent: "")
        menu.addItem(.separator())
        menu.addItem(withTitle: "Delete Row", action: #selector(GridBodyView.ctxDelRow(_:)), keyEquivalent: "")
        menu.addItem(withTitle: "Delete Column", action: #selector(GridBodyView.ctxDelCol(_:)), keyEquivalent: "")
        menu.addItem(.separator())
        menu.addItem(withTitle: "Clear Contents", action: #selector(GridBodyView.ctxClear(_:)), keyEquivalent: "")
        return menu
    }

    func invalidateGrid() {
        containerView?.invalidateGrid()
    }

    func focusGrid() {
        viewModel.formulaBarHasFocus = false
        containerView?.focusBodyView()
    }

    func clearHover() {
        hoverCell = nil
        hoverColumn = nil
        hoverRow = nil
    }

    func commitEditAndMove(direction: Direction) {
        commitAndRemoveEditor()
        clearHover()
        viewModel.moveSelection(direction: direction, extend: false)
        scrollToActiveCell()
        focusGrid()
        invalidateGrid()
    }

    func control(_ control: NSControl, textView: NSTextView, doCommandBy selector: Selector) -> Bool {
        if selector == #selector(NSResponder.insertNewline(_:)) {
            commitEditAndMove(direction: .down)
            return true
        }
        if selector == #selector(NSResponder.insertTab(_:)) {
            commitEditAndMove(direction: .right)
            return true
        }
        if selector == #selector(NSResponder.moveUp(_:)) ||
            selector == #selector(NSResponder.moveUpAndModifySelection(_:)) {
            commitEditAndMove(direction: .up)
            return true
        }
        if selector == #selector(NSResponder.moveDown(_:)) ||
            selector == #selector(NSResponder.moveDownAndModifySelection(_:)) {
            commitEditAndMove(direction: .down)
            return true
        }
        if selector == #selector(NSResponder.moveLeft(_:)) ||
            selector == #selector(NSResponder.moveLeftAndModifySelection(_:)) {
            commitEditAndMove(direction: .left)
            return true
        }
        if selector == #selector(NSResponder.moveRight(_:)) ||
            selector == #selector(NSResponder.moveRightAndModifySelection(_:)) {
            commitEditAndMove(direction: .right)
            return true
        }
        if selector == #selector(NSResponder.cancelOperation(_:)) {
            cancelAndRemoveEditor()
            focusGrid()
            return true
        }
        return false
    }

    func controlTextDidChange(_ obj: Notification) {
        guard !isApplyingEditorAttributes else { return }
        discardFormulaReferenceInsertionState()
        if let textView = obj.userInfo?["NSFieldEditor"] as? NSTextView {
            viewModel.editingText = textView.string
            activeEditor?.stringValue = textView.string
        } else if let editor = activeEditor, let fieldEditor = editor.currentEditor() as? NSTextView {
            viewModel.editingText = fieldEditor.string
            editor.stringValue = fieldEditor.string
        } else if let editor = activeEditor {
            viewModel.editingText = editor.stringValue
        }
        applyEditorTextAttributes()
        invalidateGrid()
    }
}

// MARK: - GridContainerView

final class GridContainerView: NSView {
    let coordinator: Coordinator
    let scrollView = NSScrollView()
    let bodyView: GridBodyView
    let columnHeaderView: GridColumnHeaderView
    let rowHeaderView: GridRowHeaderView
    let cornerView: GridCornerView
    private var boundsObserver: NSObjectProtocol?
    private var focusObserver: NSObjectProtocol?
    private var didRequestInitialFocus = false

    init(coordinator: Coordinator) {
        self.coordinator = coordinator
        self.bodyView = GridBodyView(coordinator: coordinator)
        self.columnHeaderView = GridColumnHeaderView(coordinator: coordinator)
        self.rowHeaderView = GridRowHeaderView(coordinator: coordinator)
        self.cornerView = GridCornerView(coordinator: coordinator)
        super.init(frame: .zero)

        coordinator.containerView = self
        coordinator.scrollView = scrollView

        wantsLayer = true
        layer?.backgroundColor = GridForgeNSColors.cellBackground.cgColor

        scrollView.hasVerticalScroller = true
        scrollView.hasHorizontalScroller = true
        scrollView.autohidesScrollers = true
        scrollView.drawsBackground = false
        scrollView.scrollerStyle = .overlay
        scrollView.usesPredominantAxisScrolling = false
        scrollView.contentView.postsBoundsChangedNotifications = true
        scrollView.documentView = bodyView

        addSubview(scrollView)
        addSubview(columnHeaderView)
        addSubview(rowHeaderView)
        addSubview(cornerView)

        boundsObserver = NotificationCenter.default.addObserver(
            forName: NSView.boundsDidChangeNotification,
            object: scrollView.contentView,
            queue: .main
        ) { [weak self] _ in
            self?.syncHeaderOffsets()
        }

        focusObserver = NotificationCenter.default.addObserver(
            forName: .gridForgeFocusGrid,
            object: nil,
            queue: .main
        ) { [weak self] _ in
            self?.focusBodyView(force: true)
        }

        updateDocumentSize()
    }

    @available(*, unavailable) required init?(coder: NSCoder) { fatalError() }

    deinit {
        if let boundsObserver {
            NotificationCenter.default.removeObserver(boundsObserver)
        }
        if let focusObserver {
            NotificationCenter.default.removeObserver(focusObserver)
        }
    }

    override var isFlipped: Bool { true }

    override func viewDidMoveToWindow() {
        super.viewDidMoveToWindow()
        guard window != nil, !didRequestInitialFocus else { return }
        didRequestInitialFocus = true
        DispatchQueue.main.async { [weak self] in
            self?.focusBodyView(force: true)
        }
    }

    override func layout() {
        super.layout()
        let rhw = coordinator.rhw
        let chh = coordinator.chh
        let width = max(0, bounds.width)
        let height = max(0, bounds.height)

        cornerView.frame = NSRect(x: 0, y: 0, width: rhw, height: chh)
        columnHeaderView.frame = NSRect(x: rhw, y: 0, width: max(0, width - rhw), height: chh)
        rowHeaderView.frame = NSRect(x: 0, y: chh, width: rhw, height: max(0, height - chh))
        scrollView.frame = NSRect(x: rhw, y: chh, width: max(0, width - rhw), height: max(0, height - chh))
        syncHeaderOffsets()
    }

    func updateDocumentSize() {
        let size = NSSize(
            width: CGFloat(coordinator.sheet.totalWidth(columns: coordinator.nCols)),
            height: CGFloat(coordinator.sheet.totalHeight(rows: coordinator.nRows))
        )
        if bodyView.frame.size != size {
            bodyView.frame = NSRect(origin: .zero, size: size)
        }
        syncHeaderOffsets()
    }

    func syncHeaderOffsets() {
        let visible = scrollView.documentVisibleRect
        columnHeaderView.scrollOffsetX = visible.minX
        rowHeaderView.scrollOffsetY = visible.minY
        columnHeaderView.needsDisplay = true
        rowHeaderView.needsDisplay = true
        cornerView.needsDisplay = true
    }

    func invalidateGrid() {
        updateDocumentSize()
        bodyView.needsDisplay = true
        columnHeaderView.needsDisplay = true
        rowHeaderView.needsDisplay = true
        cornerView.needsDisplay = true
    }

    func focusBodyView(force: Bool = false) {
        guard let window else { return }
        if !force, coordinator.viewModel.formulaBarHasFocus { return }
        window.makeFirstResponder(bodyView)
    }
}

// MARK: - Body View

final class GridBodyView: NSView {
    let coordinator: Coordinator
    private var trackingArea: NSTrackingArea?

    init(coordinator: Coordinator) {
        self.coordinator = coordinator
        super.init(frame: .zero)
        canDrawConcurrently = true
    }

    @available(*, unavailable) required init?(coder: NSCoder) { fatalError() }

    override var isFlipped: Bool { true }
    override var acceptsFirstResponder: Bool { true }

    override func updateTrackingAreas() {
        if let trackingArea {
            removeTrackingArea(trackingArea)
        }
        let area = NSTrackingArea(
            rect: bounds,
            options: [.mouseMoved, .mouseEnteredAndExited, .activeInKeyWindow, .inVisibleRect],
            owner: self,
            userInfo: nil
        )
        addTrackingArea(area)
        trackingArea = area
        super.updateTrackingAreas()
    }

    override func draw(_ dirtyRect: NSRect) {
        super.draw(dirtyRect)
        guard let ctx = NSGraphicsContext.current?.cgContext else { return }

        let co = coordinator
        let visible = visibleRect
        let sheet = co.viewModel.activeSheet
        let selection = co.viewModel.selection
        let colRange = co.visibleCols(in: visible)
        let rowRange = co.visibleRows(in: visible)
        let colXs = colRange.map { (co.xForCol($0), co.colW($0)) }
        let rowYs = rowRange.map { (co.yForRow($0), co.rowH($0)) }

        ctx.saveGState()
        ctx.clip(to: dirtyRect)

        ctx.setFillColor(GridForgeNSColors.cellBackground.cgColor)
        ctx.fill(dirtyRect)

        ctx.setStrokeColor(GridForgeNSColors.gridLine.cgColor)
        ctx.setLineWidth(0.5)
        for (x, w) in colXs {
            ctx.move(to: CGPoint(x: x + w, y: visible.minY))
            ctx.addLine(to: CGPoint(x: x + w, y: visible.maxY))
        }
        ctx.strokePath()

        for (y, h) in rowYs {
            ctx.move(to: CGPoint(x: visible.minX, y: y + h))
            ctx.addLine(to: CGPoint(x: visible.maxX, y: y + h))
        }
        ctx.strokePath()

        for (rowIndex, row) in rowRange.enumerated() {
            let (cellY, cellHeight) = rowYs[rowIndex]
            for (colIndex, column) in colRange.enumerated() {
                let (cellX, cellWidth) = colXs[colIndex]
                guard let cell = sheet.cell(at: CellAddress(column: column, row: row)), !cell.isEmpty else { continue }
                let textRect = NSRect(
                    x: cellX + GridForgeSpacing.cellPaddingH,
                    y: cellY + GridForgeSpacing.cellPaddingV,
                    width: cellWidth - GridForgeSpacing.cellPaddingH * 2,
                    height: cellHeight - GridForgeSpacing.cellPaddingV * 2
                )
                (cell.displayString as NSString).draw(in: textRect, withAttributes: co.cellAttributes(for: cell))
            }
        }

        let selectedRange = selection.selectedRange
        if !selectedRange.isSingleCell {
            let rect = rangeRect(selectedRange)
            ctx.setFillColor(GridForgeNSColors.selectedRangeFill.cgColor)
            ctx.fill(rect)
            ctx.setStrokeColor(GridForgeNSColors.selectedRangeBorder.cgColor)
            ctx.setLineWidth(1)
            ctx.stroke(rect)
        }

        if co.isEditingFormula {
            for highlight in co.formulaReferenceHighlights() {
                let rect = rangeRect(highlight.cellRange)
                guard visible.intersects(rect) else { continue }

                ctx.setFillColor(highlight.color.withAlphaComponent(0.08).cgColor)
                ctx.fill(rect)
                ctx.setStrokeColor(highlight.color.cgColor)
                ctx.setLineWidth(2)
                ctx.stroke(rect.insetBy(dx: 1, dy: 1))
            }
        }

        let active = selection.activeCell
        let activeRect = NSRect(
            x: co.xForCol(active.column),
            y: co.yForRow(active.row),
            width: co.colW(active.column),
            height: co.rowH(active.row)
        )
        ctx.setStrokeColor(GridForgeNSColors.selectedCellBorder.cgColor)
        ctx.setLineWidth(2)
        ctx.stroke(activeRect)
        ctx.setStrokeColor(NSColor.white.withAlphaComponent(0.7).cgColor)
        ctx.setLineWidth(0.5)
        ctx.stroke(activeRect.insetBy(dx: 1, dy: 1))

        if co.isResizingColumn {
            let resizeX = co.xForCol(co.resizeColumnIndex) + max(20, co.resizePreviewWidth)
            ctx.setStrokeColor(NSColor.controlAccentColor.cgColor)
            ctx.setLineWidth(1.5)
            ctx.move(to: CGPoint(x: resizeX, y: visible.minY))
            ctx.addLine(to: CGPoint(x: resizeX, y: visible.maxY))
            ctx.strokePath()
        }

        ctx.restoreGState()
    }

    private func rangeRect(_ range: CellRange) -> NSRect {
        let co = coordinator
        let x1 = co.xForCol(range.start.column)
        let y1 = co.yForRow(range.start.row)
        let x2 = co.xForCol(range.end.column) + co.colW(range.end.column)
        let y2 = co.yForRow(range.end.row) + co.rowH(range.end.row)
        return NSRect(x: x1, y: y1, width: x2 - x1, height: y2 - y1)
    }

    override func mouseDown(with event: NSEvent) {
        let location = convert(event.locationInWindow, from: nil)
        let co = coordinator
        let address = CellAddress(column: co.colAt(location.x), row: co.rowAt(location.y))

        if co.isEditingFormula {
            co.insertFormulaReference(address)
            if co.activeEditor == nil {
                co.focusGrid()
            }
            return
        }

        co.focusGrid()

        if event.clickCount == 2 {
            co.selectionDragAnchor = nil
            co.viewModel.selectCell(address)
            co.beginEditing(in: self)
            return
        }

        if co.viewModel.isEditing {
            co.commitAndRemoveEditor()
        }
        if event.modifierFlags.contains(.shift) {
            co.selectionDragAnchor = co.viewModel.selection.activeCell
            co.viewModel.selection.extendSelection(to: address)
            co.viewModel.version += 1
        } else {
            co.selectionDragAnchor = address
            co.viewModel.selectCell(address)
        }
        co.invalidateGrid()
    }

    override func mouseDragged(with event: NSEvent) {
        let location = convert(event.locationInWindow, from: nil)
        let co = coordinator
        let address = CellAddress(column: co.colAt(location.x), row: co.rowAt(location.y))

        if co.isEditingFormula {
            co.updateFormulaRangeReference(to: address)
            return
        }

        if let anchor = co.selectionDragAnchor {
            co.viewModel.selection.activeCell = anchor
        }
        co.viewModel.selection.extendSelection(to: address)
        co.viewModel.version += 1
        co.invalidateGrid()
    }

    override func mouseUp(with event: NSEvent) {
        coordinator.selectionDragAnchor = nil
        coordinator.endFormulaRangeReference()
    }

    override func mouseMoved(with event: NSEvent) {
        let location = convert(event.locationInWindow, from: nil)
        let co = coordinator
        let hover = CellAddress(column: co.colAt(location.x), row: co.rowAt(location.y))
        if co.hoverCell != hover {
            co.hoverCell = hover
            co.hoverColumn = hover.column
            co.hoverRow = hover.row
            co.invalidateGrid()
        }
    }

    override func mouseExited(with event: NSEvent) {
        let co = coordinator
        if co.hoverCell != nil || co.hoverColumn != nil || co.hoverRow != nil {
            co.hoverCell = nil
            co.hoverColumn = nil
            co.hoverRow = nil
            co.invalidateGrid()
        }
        NSCursor.arrow.set()
    }

    override func rightMouseDown(with event: NSEvent) {
        let location = convert(event.locationInWindow, from: nil)
        let co = coordinator
        let address = CellAddress(column: co.colAt(location.x), row: co.rowAt(location.y))
        if !co.viewModel.selection.selectedRange.contains(address) {
            co.viewModel.selectCell(address)
            co.invalidateGrid()
        }
        NSMenu.popUpContextMenu(co.buildContextMenu(), with: event, for: self)
    }

    @objc func ctxCut(_ sender: Any?) {
        coordinator.viewModel.copy()
        coordinator.viewModel.deleteSelectedCells()
        coordinator.invalidateGrid()
    }

    @objc func ctxCopy(_ sender: Any?) {
        coordinator.viewModel.copy()
    }

    @objc func ctxPaste(_ sender: Any?) {
        coordinator.viewModel.paste()
        coordinator.invalidateGrid()
    }

    @objc func ctxInsRowAbove(_ sender: Any?) {
        coordinator.viewModel.insertRow()
        coordinator.invalidateGrid()
    }

    @objc func ctxInsRowBelow(_ sender: Any?) {
        let vm = coordinator.viewModel
        vm.selectCell(CellAddress(column: vm.activeCell.column, row: vm.activeCell.row + 1))
        vm.insertRow()
        coordinator.invalidateGrid()
    }

    @objc func ctxInsColLeft(_ sender: Any?) {
        coordinator.viewModel.insertColumn()
        coordinator.invalidateGrid()
    }

    @objc func ctxInsColRight(_ sender: Any?) {
        let vm = coordinator.viewModel
        vm.selectCell(CellAddress(column: vm.activeCell.column + 1, row: vm.activeCell.row))
        vm.insertColumn()
        coordinator.invalidateGrid()
    }

    @objc func ctxDelRow(_ sender: Any?) {
        coordinator.viewModel.deleteRow()
        coordinator.invalidateGrid()
    }

    @objc func ctxDelCol(_ sender: Any?) {
        coordinator.viewModel.deleteColumn()
        coordinator.invalidateGrid()
    }

    @objc func ctxClear(_ sender: Any?) {
        coordinator.viewModel.deleteSelectedCells()
        coordinator.invalidateGrid()
    }

    override func keyDown(with event: NSEvent) {
        let vm = coordinator.viewModel
        let shift = event.modifierFlags.contains(.shift)
        let command = event.modifierFlags.contains(.command)
        vm.formulaBarHasFocus = false

        if command {
            switch event.charactersIgnoringModifiers {
            case "c":
                vm.copy()
                return
            case "x":
                vm.copy()
                vm.deleteSelectedCells()
                coordinator.invalidateGrid()
                return
            case "v":
                vm.paste()
                coordinator.invalidateGrid()
                return
            case "z":
                shift ? vm.redo() : vm.undo()
                coordinator.invalidateGrid()
                return
            case "a":
                vm.selectAll()
                coordinator.invalidateGrid()
                return
            case "b":
                vm.toggleBold()
                coordinator.invalidateGrid()
                return
            case "i":
                vm.toggleItalic()
                coordinator.invalidateGrid()
                return
            default:
                break
            }
            if event.keyCode == 115 {
                vm.selectCell(CellAddress(column: 0, row: 0))
                coordinator.scrollToActiveCell()
                coordinator.invalidateGrid()
                return
            }
        }

        switch event.keyCode {
        case 126:
            coordinator.clearHover()
            vm.moveSelection(direction: .up, extend: shift)
            coordinator.scrollToActiveCell()
            coordinator.invalidateGrid()
            return
        case 125:
            coordinator.clearHover()
            vm.moveSelection(direction: .down, extend: shift)
            coordinator.scrollToActiveCell()
            coordinator.invalidateGrid()
            return
        case 123:
            coordinator.clearHover()
            vm.moveSelection(direction: .left, extend: shift)
            coordinator.scrollToActiveCell()
            coordinator.invalidateGrid()
            return
        case 124:
            coordinator.clearHover()
            vm.moveSelection(direction: .right, extend: shift)
            coordinator.scrollToActiveCell()
            coordinator.invalidateGrid()
            return
        case 36:
            if vm.isEditing {
                coordinator.commitAndRemoveEditor()
                vm.moveSelection(direction: .down, extend: false)
            } else {
                coordinator.beginEditing(in: self)
            }
            coordinator.invalidateGrid()
            return
        case 48:
            if vm.isEditing {
                coordinator.commitAndRemoveEditor()
            }
            vm.moveSelection(direction: .right, extend: false)
            coordinator.scrollToActiveCell()
            coordinator.invalidateGrid()
            return
        case 53:
            if vm.isEditing {
                coordinator.cancelAndRemoveEditor()
            }
            return
        case 51, 117:
            if !vm.isEditing {
                vm.deleteSelectedCells()
                coordinator.invalidateGrid()
            }
            return
        case 120:
            if !vm.isEditing {
                coordinator.beginEditing(in: self)
                coordinator.invalidateGrid()
            }
            return
        case 116:
            pageMove(direction: .up, shift: shift)
            return
        case 121:
            pageMove(direction: .down, shift: shift)
            return
        case 115:
            vm.selectCell(CellAddress(column: 0, row: vm.activeCell.row))
            coordinator.scrollToActiveCell()
            coordinator.invalidateGrid()
            return
        case 119:
            let lastCol = vm.activeSheet.usedRange?.end.column ?? 0
            vm.selectCell(CellAddress(column: lastCol, row: vm.activeCell.row))
            coordinator.scrollToActiveCell()
            coordinator.invalidateGrid()
            return
        default:
            break
        }

        if let chars = event.characters, !chars.isEmpty, !command, coordinator.isEditingFormula {
            coordinator.insertTextIntoFormulaEditor(chars)
            return
        }

        if let chars = event.characters, !chars.isEmpty, !command, let first = chars.first, first.isPrintable {
            coordinator.beginEditing(in: self, withText: String(first))
            return
        }

        super.keyDown(with: event)
    }

    private func pageMove(direction: Direction, shift: Bool) {
        guard let scrollView = coordinator.scrollView else { return }
        let rows = max(1, Int(scrollView.documentVisibleRect.height / CGFloat(Worksheet.defaultRowHeight)))
        for _ in 0..<rows {
            coordinator.viewModel.moveSelection(direction: direction, extend: shift)
        }
        coordinator.scrollToActiveCell()
        coordinator.invalidateGrid()
    }
}

// MARK: - Header Views

final class GridColumnHeaderView: NSView {
    let coordinator: Coordinator
    var scrollOffsetX: CGFloat = 0
    private var trackingArea: NSTrackingArea?

    init(coordinator: Coordinator) {
        self.coordinator = coordinator
        super.init(frame: .zero)
        canDrawConcurrently = true
    }

    @available(*, unavailable) required init?(coder: NSCoder) { fatalError() }

    override var isFlipped: Bool { true }

    override func updateTrackingAreas() {
        if let trackingArea {
            removeTrackingArea(trackingArea)
        }
        let area = NSTrackingArea(
            rect: bounds,
            options: [.mouseMoved, .mouseEnteredAndExited, .activeInKeyWindow, .inVisibleRect],
            owner: self,
            userInfo: nil
        )
        addTrackingArea(area)
        trackingArea = area
        super.updateTrackingAreas()
    }

    override func draw(_ dirtyRect: NSRect) {
        super.draw(dirtyRect)
        guard let ctx = NSGraphicsContext.current?.cgContext else { return }
        let co = coordinator
        let selectedRange = co.viewModel.selection.selectedRange
        let visible = NSRect(x: scrollOffsetX, y: 0, width: bounds.width, height: bounds.height)
        let colRange = co.visibleCols(in: visible)

        ctx.saveGState()
        ctx.clip(to: dirtyRect)
        ctx.setFillColor(GridForgeNSColors.headerBackground.cgColor)
        ctx.fill(bounds)
        drawHeaderGradient(ctx: ctx, rect: bounds, vertical: true)

        ctx.setStrokeColor(GridForgeNSColors.headerBorder.cgColor)
        ctx.setLineWidth(1)
        ctx.move(to: CGPoint(x: 0, y: bounds.maxY))
        ctx.addLine(to: CGPoint(x: bounds.maxX, y: bounds.maxY))
        ctx.strokePath()

        for column in colRange {
            let x = co.xForCol(column) - scrollOffsetX
            let width = co.colW(column)
            let rect = NSRect(x: x, y: 0, width: width, height: bounds.height)
            let isActive = column >= selectedRange.start.column && column <= selectedRange.end.column
            let isHovered = co.hoverColumn == column

            if isActive {
                ctx.setFillColor(GridForgeNSColors.activeHeaderHighlight.cgColor)
                ctx.fill(rect)
            } else if isHovered {
                ctx.setFillColor(GridForgeNSColors.headerHover.cgColor)
                ctx.fill(rect)
            }

            ctx.setStrokeColor(GridForgeNSColors.gridLine.cgColor)
            ctx.setLineWidth(0.5)
            ctx.move(to: CGPoint(x: rect.maxX, y: rect.minY))
            ctx.addLine(to: CGPoint(x: rect.maxX, y: rect.maxY))
            ctx.strokePath()

            let textRect = NSRect(x: rect.minX, y: rect.minY + 6, width: rect.width, height: max(0, rect.height - 6))
            (CellAddress(column: column, row: 0).columnLetter as NSString)
                .draw(in: textRect, withAttributes: isActive ? co.activeHeaderAttrs : co.headerAttrs)
        }

        if co.isResizingColumn {
            let x = co.xForCol(co.resizeColumnIndex) + max(20, co.resizePreviewWidth) - scrollOffsetX
            ctx.setStrokeColor(NSColor.controlAccentColor.cgColor)
            ctx.setLineWidth(1.5)
            ctx.move(to: CGPoint(x: x, y: bounds.minY))
            ctx.addLine(to: CGPoint(x: x, y: bounds.maxY))
            ctx.strokePath()
        }

        ctx.restoreGState()
    }

    override func mouseDown(with event: NSEvent) {
        window?.makeFirstResponder(coordinator.containerView?.bodyView)
        let location = convert(event.locationInWindow, from: nil)
        let contentX = location.x + scrollOffsetX
        guard let column = coordinator.colBorderHit(atContentX: contentX) else { return }
        coordinator.isResizingColumn = true
        coordinator.resizeColumnIndex = column
        coordinator.resizeStartX = contentX
        coordinator.resizeStartWidth = coordinator.colW(column)
        coordinator.resizePreviewWidth = coordinator.resizeStartWidth
        coordinator.invalidateGrid()
    }

    override func mouseDragged(with event: NSEvent) {
        guard coordinator.isResizingColumn else { return }
        let location = convert(event.locationInWindow, from: nil)
        let contentX = location.x + scrollOffsetX
        coordinator.resizePreviewWidth = max(20, coordinator.resizeStartWidth + contentX - coordinator.resizeStartX)
        coordinator.invalidateGrid()
    }

    override func mouseUp(with event: NSEvent) {
        guard coordinator.isResizingColumn else { return }
        coordinator.viewModel.setColumnWidth(coordinator.resizeColumnIndex, Double(coordinator.resizePreviewWidth))
        coordinator.isResizingColumn = false
        coordinator.containerView?.updateDocumentSize()
        coordinator.invalidateGrid()
    }

    override func mouseMoved(with event: NSEvent) {
        let location = convert(event.locationInWindow, from: nil)
        let contentX = location.x + scrollOffsetX
        NSCursor.resizeLeftRight.set(coordinator.colBorderHit(atContentX: contentX) != nil)
        let column = coordinator.colAt(contentX)
        if coordinator.hoverColumn != column {
            coordinator.hoverColumn = column
            coordinator.columnHeaderView?.needsDisplay = true
        }
    }

    override func mouseExited(with event: NSEvent) {
        coordinator.hoverColumn = nil
        needsDisplay = true
        NSCursor.arrow.set()
    }
}

final class GridRowHeaderView: NSView {
    let coordinator: Coordinator
    var scrollOffsetY: CGFloat = 0
    private var trackingArea: NSTrackingArea?

    init(coordinator: Coordinator) {
        self.coordinator = coordinator
        super.init(frame: .zero)
        canDrawConcurrently = true
    }

    @available(*, unavailable) required init?(coder: NSCoder) { fatalError() }

    override var isFlipped: Bool { true }

    override func updateTrackingAreas() {
        if let trackingArea {
            removeTrackingArea(trackingArea)
        }
        let area = NSTrackingArea(
            rect: bounds,
            options: [.mouseMoved, .mouseEnteredAndExited, .activeInKeyWindow, .inVisibleRect],
            owner: self,
            userInfo: nil
        )
        addTrackingArea(area)
        trackingArea = area
        super.updateTrackingAreas()
    }

    override func draw(_ dirtyRect: NSRect) {
        super.draw(dirtyRect)
        guard let ctx = NSGraphicsContext.current?.cgContext else { return }
        let co = coordinator
        let selectedRange = co.viewModel.selection.selectedRange
        let visible = NSRect(x: 0, y: scrollOffsetY, width: bounds.width, height: bounds.height)
        let rowRange = co.visibleRows(in: visible)

        ctx.saveGState()
        ctx.clip(to: dirtyRect)
        ctx.setFillColor(GridForgeNSColors.headerBackground.cgColor)
        ctx.fill(bounds)
        drawHeaderGradient(ctx: ctx, rect: bounds, vertical: false)

        ctx.setStrokeColor(GridForgeNSColors.headerBorder.cgColor)
        ctx.setLineWidth(1)
        ctx.move(to: CGPoint(x: bounds.maxX, y: bounds.minY))
        ctx.addLine(to: CGPoint(x: bounds.maxX, y: bounds.maxY))
        ctx.strokePath()

        for row in rowRange {
            let y = co.yForRow(row) - scrollOffsetY
            let height = co.rowH(row)
            let rect = NSRect(x: 0, y: y, width: bounds.width, height: height)
            let isActive = row >= selectedRange.start.row && row <= selectedRange.end.row
            let isHovered = co.hoverRow == row

            if isActive {
                ctx.setFillColor(GridForgeNSColors.activeHeaderHighlight.cgColor)
                ctx.fill(rect)
            } else if isHovered {
                ctx.setFillColor(GridForgeNSColors.headerHover.cgColor)
                ctx.fill(rect)
            }

            ctx.setStrokeColor(GridForgeNSColors.gridLine.cgColor)
            ctx.setLineWidth(0.5)
            ctx.move(to: CGPoint(x: rect.minX, y: rect.maxY))
            ctx.addLine(to: CGPoint(x: rect.maxX, y: rect.maxY))
            ctx.strokePath()

            let textRect = NSRect(x: 0, y: y + 4, width: bounds.width, height: max(0, height - 4))
            ("\(row + 1)" as NSString).draw(in: textRect, withAttributes: isActive ? co.activeHeaderAttrs : co.headerAttrs)
        }

        ctx.restoreGState()
    }

    override func mouseMoved(with event: NSEvent) {
        let location = convert(event.locationInWindow, from: nil)
        let row = coordinator.rowAt(location.y + scrollOffsetY)
        if coordinator.hoverRow != row {
            coordinator.hoverRow = row
            needsDisplay = true
        }
    }

    override func mouseExited(with event: NSEvent) {
        coordinator.hoverRow = nil
        needsDisplay = true
    }
}

final class GridCornerView: NSView {
    let coordinator: Coordinator

    init(coordinator: Coordinator) {
        self.coordinator = coordinator
        super.init(frame: .zero)
    }

    @available(*, unavailable) required init?(coder: NSCoder) { fatalError() }

    override var isFlipped: Bool { true }

    override func draw(_ dirtyRect: NSRect) {
        super.draw(dirtyRect)
        guard let ctx = NSGraphicsContext.current?.cgContext else { return }

        ctx.setFillColor(GridForgeNSColors.cornerBackground.cgColor)
        ctx.fill(bounds)
        ctx.setStrokeColor(GridForgeNSColors.headerBorder.cgColor)
        ctx.setLineWidth(1)
        ctx.move(to: CGPoint(x: bounds.maxX, y: bounds.minY))
        ctx.addLine(to: CGPoint(x: bounds.maxX, y: bounds.maxY))
        ctx.move(to: CGPoint(x: bounds.minX, y: bounds.maxY))
        ctx.addLine(to: CGPoint(x: bounds.maxX, y: bounds.maxY))
        ctx.strokePath()
    }
}

// MARK: - Drawing Helpers

private func drawHeaderGradient(ctx: CGContext, rect: NSRect, vertical: Bool) {
    let top = NSColor.white.withAlphaComponent(0.08).cgColor
    let bottom = NSColor.black.withAlphaComponent(0.04).cgColor
    guard let gradient = CGGradient(
        colorsSpace: CGColorSpaceCreateDeviceRGB(),
        colors: [top, bottom] as CFArray,
        locations: [0, 1]
    ) else { return }

    ctx.saveGState()
    ctx.clip(to: rect)
    if vertical {
        ctx.drawLinearGradient(
            gradient,
            start: CGPoint(x: rect.midX, y: rect.minY),
            end: CGPoint(x: rect.midX, y: rect.maxY),
            options: []
        )
    } else {
        ctx.drawLinearGradient(
            gradient,
            start: CGPoint(x: rect.minX, y: rect.midY),
            end: CGPoint(x: rect.maxX, y: rect.midY),
            options: []
        )
    }
    ctx.restoreGState()
}

private extension Coordinator {
    var columnHeaderView: GridColumnHeaderView? {
        containerView?.columnHeaderView
    }
}

private extension NSCursor {
    func set(_ condition: Bool) {
        condition ? set() : NSCursor.arrow.set()
    }
}

private extension Character {
    var isPrintable: Bool {
        guard let scalar = unicodeScalars.first else { return false }
        return scalar.value >= 32 && scalar.value < 127 || scalar.properties.isAlphabetic || scalar.properties.isEmoji
    }
}
