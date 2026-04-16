import SwiftUI
import AppKit
import GridForgeCore

// MARK: - SpreadsheetGridView (NSViewRepresentable)

struct SpreadsheetGridView: NSViewRepresentable {
    @ObservedObject var viewModel: WorkbookViewModel

    func makeCoordinator() -> Coordinator { Coordinator(viewModel: viewModel) }

    func makeNSView(context: Context) -> NSScrollView {
        let sv = NSScrollView()
        sv.hasVerticalScroller = true
        sv.hasHorizontalScroller = true
        sv.autohidesScrollers = true
        sv.drawsBackground = false
        sv.scrollerStyle = .overlay
        sv.usesPredominantAxisScrolling = false
        let gv = GridDrawingView(coordinator: context.coordinator)
        gv.frame = NSRect(x: 0, y: 0, width: contentWidth(), height: contentHeight())
        sv.documentView = gv
        context.coordinator.scrollView = sv
        return sv
    }

    func updateNSView(_ sv: NSScrollView, context: Context) {
        context.coordinator.viewModel = viewModel
        context.coordinator.scrollView = sv
        if let gv = sv.documentView as? GridDrawingView {
            let size = NSSize(width: contentWidth(), height: contentHeight())
            if gv.frame.size != size { gv.frame = NSRect(origin: .zero, size: size) }
            gv.needsDisplay = true
        }
    }

    private func contentWidth() -> CGFloat {
        GridForgeSpacing.rowHeaderWidth + CGFloat(viewModel.activeSheet.totalWidth(columns: viewModel.displayColumns))
    }
    private func contentHeight() -> CGFloat {
        GridForgeSpacing.columnHeaderHeight + CGFloat(viewModel.activeSheet.totalHeight(rows: viewModel.displayRows))
    }
}

// MARK: - Coordinator

final class Coordinator: NSObject, NSTextFieldDelegate {
    var viewModel: WorkbookViewModel
    weak var activeEditor: NSTextField?
    weak var scrollView: NSScrollView?
    var hoverCell: CellAddress?
    var isResizingColumn = false
    var resizeColumnIndex = -1
    var resizeStartX: CGFloat = 0
    var resizeStartWidth: CGFloat = 0

    // Cached font attribute dictionaries
    let cellAttrsNormal: [NSAttributedString.Key: Any]
    let cellAttrsBold: [NSAttributedString.Key: Any]
    let cellAttrsItalic: [NSAttributedString.Key: Any]
    let cellAttrsBoldItalic: [NSAttributedString.Key: Any]
    let headerAttrs: [NSAttributedString.Key: Any]
    let activeHeaderAttrs: [NSAttributedString.Key: Any]

    init(viewModel: WorkbookViewModel) {
        self.viewModel = viewModel
        let centerPara = NSMutableParagraphStyle()
        centerPara.alignment = .center; centerPara.lineBreakMode = .byClipping
        headerAttrs = [.font: GridForgeNSFonts.headerFont, .foregroundColor: GridForgeNSColors.headerText, .paragraphStyle: centerPara]
        activeHeaderAttrs = [.font: NSFont.systemFont(ofSize: 11, weight: .bold), .foregroundColor: NSColor.controlAccentColor, .paragraphStyle: centerPara]
        let clipPara = NSMutableParagraphStyle(); clipPara.lineBreakMode = .byClipping
        cellAttrsNormal = [.font: GridForgeNSFonts.cellFont, .foregroundColor: GridForgeNSColors.cellText, .paragraphStyle: clipPara]
        cellAttrsBold = [.font: GridForgeNSFonts.cellFontBold, .foregroundColor: GridForgeNSColors.cellText, .paragraphStyle: clipPara]
        cellAttrsItalic = [.font: GridForgeNSFonts.cellFontItalic, .foregroundColor: GridForgeNSColors.cellText, .paragraphStyle: clipPara]
        cellAttrsBoldItalic = [.font: GridForgeNSFonts.cellFontBoldItalic, .foregroundColor: GridForgeNSColors.cellText, .paragraphStyle: clipPara]
        super.init()
    }

    // MARK: Geometry (cached offsets + binary search)

    var rhw: CGFloat { GridForgeSpacing.rowHeaderWidth }
    var chh: CGFloat { GridForgeSpacing.columnHeaderHeight }
    var sheet: Worksheet { viewModel.activeSheet }
    var nCols: Int { viewModel.displayColumns }
    var nRows: Int { viewModel.displayRows }

    func xForCol(_ c: Int) -> CGFloat { rhw + CGFloat(sheet.xOffset(for: c, totalColumns: nCols)) }
    func yForRow(_ r: Int) -> CGFloat { chh + CGFloat(sheet.yOffset(for: r, totalRows: nRows)) }
    func colW(_ c: Int) -> CGFloat { CGFloat(sheet.columnWidth(for: c)) }
    func rowH(_ r: Int) -> CGFloat { CGFloat(sheet.rowHeight(for: r)) }
    func colAt(_ x: CGFloat) -> Int { sheet.columnAt(x: Double(x - rhw), totalColumns: nCols) }
    func rowAt(_ y: CGFloat) -> Int { sheet.rowAt(y: Double(y - chh), totalRows: nRows) }
    func visibleCols(in r: NSRect) -> ClosedRange<Int> { colAt(r.minX)...colAt(r.maxX) }
    func visibleRows(in r: NSRect) -> ClosedRange<Int> { rowAt(r.minY)...rowAt(r.maxY) }

    func colBorderHit(at x: CGFloat, tol: CGFloat = 4) -> Int? {
        guard let sv = scrollView else { return nil }
        for c in visibleCols(in: sv.documentVisibleRect) {
            if abs(x - (xForCol(c) + colW(c))) <= tol { return c }
        }
        return nil
    }

    func cellAttributes(for cell: Cell) -> [NSAttributedString.Key: Any] {
        var attrs = (cell.formatting.bold && cell.formatting.italic) ? cellAttrsBoldItalic
            : cell.formatting.bold ? cellAttrsBold
            : cell.formatting.italic ? cellAttrsItalic : cellAttrsNormal
        if cell.value.isError { attrs[.foregroundColor] = GridForgeNSColors.errorText }
        let p = NSMutableParagraphStyle(); p.lineBreakMode = .byClipping
        switch cell.formatting.alignment {
        case .left:    p.alignment = .left
        case .center:  p.alignment = .center
        case .right:   p.alignment = .right
        case .general: if case .number = cell.value { p.alignment = .right } else { p.alignment = .left }
        }
        attrs[.paragraphStyle] = p
        return attrs
    }

    // MARK: Editing

    func beginEditing(in gv: GridDrawingView, withText text: String? = nil) {
        removeEditor()
        let a = viewModel.activeCell
        viewModel.startEditing(withText: text)
        let rect = NSRect(x: xForCol(a.column) - 1, y: yForRow(a.row) - 1, width: colW(a.column) + 2, height: rowH(a.row) + 2)
        let ed = NSTextField(frame: rect)
        ed.font = GridForgeNSFonts.editorFont
        ed.isBordered = false; ed.isBezeled = false; ed.focusRingType = .none
        ed.drawsBackground = true; ed.backgroundColor = GridForgeNSColors.editorBackground
        ed.wantsLayer = true
        ed.layer?.borderColor = GridForgeNSColors.editorBorder.cgColor
        ed.layer?.borderWidth = 2; ed.layer?.cornerRadius = 1
        ed.layer?.shadowColor = GridForgeNSColors.editorShadow.cgColor
        ed.layer?.shadowOffset = CGSize(width: 0, height: -1)
        ed.layer?.shadowRadius = 3; ed.layer?.shadowOpacity = 1
        ed.delegate = self; ed.stringValue = viewModel.editingText
        ed.isEditable = true; ed.isSelectable = true
        gv.addSubview(ed); ed.selectText(nil)
        if let text = text, !text.isEmpty, let fe = ed.currentEditor() {
            fe.selectedRange = NSRange(location: text.count, length: 0)
        }
        activeEditor = ed
    }

    func removeEditor() { activeEditor?.removeFromSuperview(); activeEditor = nil }

    func commitAndRemoveEditor() {
        if let ed = activeEditor { viewModel.editingText = ed.stringValue }
        viewModel.commitEdit(); removeEditor()
    }

    func cancelAndRemoveEditor() { viewModel.cancelEdit(); removeEditor() }

    func scrollToActiveCell() {
        guard let sv = scrollView else { return }
        let c = viewModel.activeCell
        let cellRect = NSRect(x: xForCol(c.column), y: yForRow(c.row), width: colW(c.column), height: rowH(c.row))
        let vis = sv.documentVisibleRect
        var o = vis.origin
        if cellRect.maxX > vis.maxX { o.x = cellRect.maxX - vis.width }
        if cellRect.minX < vis.minX + rhw { o.x = cellRect.minX - rhw }
        if cellRect.maxY > vis.maxY { o.y = cellRect.maxY - vis.height }
        if cellRect.minY < vis.minY + chh { o.y = cellRect.minY - chh }
        o.x = max(0, o.x); o.y = max(0, o.y)
        if o != vis.origin { sv.contentView.scroll(to: o); sv.reflectScrolledClipView(sv.contentView) }
    }

    // MARK: Context menu

    func buildContextMenu() -> NSMenu {
        let m = NSMenu()
        m.addItem(withTitle: "Cut", action: #selector(GridDrawingView.ctxCut(_:)), keyEquivalent: "x")
        m.addItem(withTitle: "Copy", action: #selector(GridDrawingView.ctxCopy(_:)), keyEquivalent: "c")
        m.addItem(withTitle: "Paste", action: #selector(GridDrawingView.ctxPaste(_:)), keyEquivalent: "v")
        m.addItem(.separator())
        m.addItem(withTitle: "Insert Row Above", action: #selector(GridDrawingView.ctxInsRowAbove(_:)), keyEquivalent: "")
        m.addItem(withTitle: "Insert Row Below", action: #selector(GridDrawingView.ctxInsRowBelow(_:)), keyEquivalent: "")
        m.addItem(withTitle: "Insert Column Left", action: #selector(GridDrawingView.ctxInsColLeft(_:)), keyEquivalent: "")
        m.addItem(withTitle: "Insert Column Right", action: #selector(GridDrawingView.ctxInsColRight(_:)), keyEquivalent: "")
        m.addItem(.separator())
        m.addItem(withTitle: "Delete Row", action: #selector(GridDrawingView.ctxDelRow(_:)), keyEquivalent: "")
        m.addItem(withTitle: "Delete Column", action: #selector(GridDrawingView.ctxDelCol(_:)), keyEquivalent: "")
        m.addItem(.separator())
        m.addItem(withTitle: "Clear Contents", action: #selector(GridDrawingView.ctxClear(_:)), keyEquivalent: "")
        return m
    }

    // MARK: NSTextFieldDelegate

    func control(_ control: NSControl, textView: NSTextView, doCommandBy sel: Selector) -> Bool {
        if sel == #selector(NSResponder.insertNewline(_:)) { commitAndRemoveEditor(); viewModel.moveSelection(direction: .down, extend: false); return true }
        if sel == #selector(NSResponder.insertTab(_:)) { commitAndRemoveEditor(); viewModel.moveSelection(direction: .right, extend: false); return true }
        if sel == #selector(NSResponder.cancelOperation(_:)) { cancelAndRemoveEditor(); return true }
        return false
    }

    func controlTextDidChange(_ obj: Notification) {
        if let ed = activeEditor { viewModel.editingText = ed.stringValue }
    }
}

// MARK: - GridDrawingView

final class GridDrawingView: NSView {
    let coordinator: Coordinator
    private var trackingArea: NSTrackingArea?

    init(coordinator: Coordinator) {
        self.coordinator = coordinator
        super.init(frame: .zero)
    }

    @available(*, unavailable) required init?(coder: NSCoder) { fatalError() }
    override var isFlipped: Bool { true }
    override var acceptsFirstResponder: Bool { true }

    override func updateTrackingAreas() {
        if let ta = trackingArea { removeTrackingArea(ta) }
        let ta = NSTrackingArea(rect: bounds, options: [.mouseMoved, .mouseEnteredAndExited, .activeInKeyWindow, .inVisibleRect], owner: self, userInfo: nil)
        addTrackingArea(ta); trackingArea = ta
        super.updateTrackingAreas()
    }

    // MARK: Drawing

    override func draw(_ dirtyRect: NSRect) {
        super.draw(dirtyRect)
        guard let ctx = NSGraphicsContext.current?.cgContext else { return }
        let co = coordinator
        let vis = visibleRect
        let vm = co.viewModel; let sheet = vm.activeSheet; let sel = vm.selection
        let rhw = co.rhw; let chh = co.chh
        let colRange = co.visibleCols(in: vis); let rowRange = co.visibleRows(in: vis)

        // Pre-compute positions for visible cells
        let colXs = colRange.map { (co.xForCol($0), co.colW($0)) }
        let rowYs = rowRange.map { (co.yForRow($0), co.rowH($0)) }

        ctx.saveGState(); ctx.clip(to: dirtyRect)

        // 1. Cell background
        ctx.setFillColor(GridForgeNSColors.cellBackground.cgColor)
        ctx.fill(NSRect(x: max(vis.minX, rhw), y: max(vis.minY, chh),
                        width: vis.maxX - max(vis.minX, rhw), height: vis.maxY - max(vis.minY, chh)))

        // 2. Batch grid lines — vertical then horizontal
        ctx.setStrokeColor(GridForgeNSColors.gridLine.cgColor); ctx.setLineWidth(0.5)
        for (x, w) in colXs { ctx.move(to: CGPoint(x: x + w, y: max(vis.minY, chh))); ctx.addLine(to: CGPoint(x: x + w, y: vis.maxY)) }
        ctx.strokePath()
        for (y, h) in rowYs { ctx.move(to: CGPoint(x: max(vis.minX, rhw), y: y + h)); ctx.addLine(to: CGPoint(x: vis.maxX, y: y + h)) }
        ctx.strokePath()

        // 3. Cell text
        for (ri, r) in rowRange.enumerated() {
            let (cy, ch) = rowYs[ri]
            for (ci, c) in colRange.enumerated() {
                let (cx, cw) = colXs[ci]
                guard let cell = sheet.cell(at: CellAddress(column: c, row: r)), !cell.isEmpty else { continue }
                let tr = NSRect(x: cx + GridForgeSpacing.cellPaddingH, y: cy + GridForgeSpacing.cellPaddingV,
                                width: cw - GridForgeSpacing.cellPaddingH * 2, height: ch - GridForgeSpacing.cellPaddingV * 2)
                (cell.displayString as NSString).draw(in: tr, withAttributes: co.cellAttributes(for: cell))
            }
        }

        // 4. Selection range fill + border
        let selR = sel.selectedRange
        if !selR.isSingleCell {
            let rr = rangeRect(selR)
            ctx.setFillColor(GridForgeNSColors.selectedRangeFill.cgColor); ctx.fill(rr)
            ctx.setStrokeColor(GridForgeNSColors.selectedRangeBorder.cgColor); ctx.setLineWidth(1); ctx.stroke(rr)
        }

        // 5. Active cell border (2px accent + 0.5px inner white)
        let ac = sel.activeCell
        let acR = NSRect(x: co.xForCol(ac.column), y: co.yForRow(ac.row), width: co.colW(ac.column), height: co.rowH(ac.row))
        ctx.setStrokeColor(GridForgeNSColors.selectedCellBorder.cgColor); ctx.setLineWidth(2); ctx.stroke(acR)
        ctx.setStrokeColor(NSColor.white.withAlphaComponent(0.7).cgColor); ctx.setLineWidth(0.5); ctx.stroke(acR.insetBy(dx: 1, dy: 1))

        // 6. Hover highlight
        if !vm.isEditing, let hv = co.hoverCell, hv != ac {
            ctx.setFillColor(GridForgeNSColors.cellHover.cgColor)
            ctx.fill(NSRect(x: co.xForCol(hv.column), y: co.yForRow(hv.row), width: co.colW(hv.column), height: co.rowH(hv.row)))
        }

        // 7. Column headers (frozen)
        let hy = vis.minY
        let chRect = NSRect(x: vis.minX, y: hy, width: vis.width, height: chh)
        ctx.setFillColor(GridForgeNSColors.headerBackground.cgColor); ctx.fill(chRect)
        drawHeaderGradient(ctx: ctx, rect: chRect, vertical: true)
        ctx.setStrokeColor(GridForgeNSColors.headerBorder.cgColor); ctx.setLineWidth(1)
        ctx.move(to: CGPoint(x: vis.minX, y: hy + chh)); ctx.addLine(to: CGPoint(x: vis.maxX, y: hy + chh)); ctx.strokePath()

        for (ci, c) in colRange.enumerated() {
            let (cx, cw) = colXs[ci]; let cr = NSRect(x: cx, y: hy, width: cw, height: chh)
            let isAct = c >= selR.start.column && c <= selR.end.column
            let isHov = co.hoverCell?.column == c
            if isAct { ctx.setFillColor(GridForgeNSColors.activeHeaderHighlight.cgColor); ctx.fill(cr) }
            else if isHov { ctx.setFillColor(GridForgeNSColors.headerHover.cgColor); ctx.fill(cr) }
            ctx.setStrokeColor(GridForgeNSColors.gridLine.cgColor); ctx.setLineWidth(0.5)
            ctx.move(to: CGPoint(x: cr.maxX, y: hy)); ctx.addLine(to: CGPoint(x: cr.maxX, y: hy + chh)); ctx.strokePath()
            let ltr = NSRect(x: cx, y: hy + 6, width: cw, height: chh - 6)
            (CellAddress(column: c, row: 0).columnLetter as NSString).draw(in: ltr, withAttributes: isAct ? co.activeHeaderAttrs : co.headerAttrs)
        }

        // 8. Row headers (frozen)
        let hx = vis.minX
        let rhRect = NSRect(x: hx, y: vis.minY, width: rhw, height: vis.height)
        ctx.setFillColor(GridForgeNSColors.headerBackground.cgColor); ctx.fill(rhRect)
        drawHeaderGradient(ctx: ctx, rect: rhRect, vertical: false)
        ctx.setStrokeColor(GridForgeNSColors.headerBorder.cgColor); ctx.setLineWidth(1)
        ctx.move(to: CGPoint(x: hx + rhw, y: vis.minY)); ctx.addLine(to: CGPoint(x: hx + rhw, y: vis.maxY)); ctx.strokePath()

        for (ri, r) in rowRange.enumerated() {
            let (ry, rh) = rowYs[ri]; let rr = NSRect(x: hx, y: ry, width: rhw, height: rh)
            let isAct = r >= selR.start.row && r <= selR.end.row
            let isHov = co.hoverCell?.row == r
            if isAct { ctx.setFillColor(GridForgeNSColors.activeHeaderHighlight.cgColor); ctx.fill(rr) }
            else if isHov { ctx.setFillColor(GridForgeNSColors.headerHover.cgColor); ctx.fill(rr) }
            ctx.setStrokeColor(GridForgeNSColors.gridLine.cgColor); ctx.setLineWidth(0.5)
            ctx.move(to: CGPoint(x: hx, y: rr.maxY)); ctx.addLine(to: CGPoint(x: hx + rhw, y: rr.maxY)); ctx.strokePath()
            let tr = NSRect(x: hx, y: ry + 4, width: rhw, height: rh - 4)
            ("\(r + 1)" as NSString).draw(in: tr, withAttributes: isAct ? co.activeHeaderAttrs : co.headerAttrs)
        }

        // 9. Corner
        let corner = NSRect(x: vis.minX, y: vis.minY, width: rhw, height: chh)
        ctx.setFillColor(GridForgeNSColors.cornerBackground.cgColor); ctx.fill(corner)
        ctx.setStrokeColor(GridForgeNSColors.headerBorder.cgColor); ctx.setLineWidth(1)
        ctx.move(to: CGPoint(x: corner.maxX, y: corner.minY)); ctx.addLine(to: CGPoint(x: corner.maxX, y: corner.maxY))
        ctx.move(to: CGPoint(x: corner.minX, y: corner.maxY)); ctx.addLine(to: CGPoint(x: corner.maxX, y: corner.maxY)); ctx.strokePath()

        // 10. Resize handle overlay
        if co.isResizingColumn {
            let rx = co.xForCol(co.resizeColumnIndex) + co.colW(co.resizeColumnIndex)
            ctx.setStrokeColor(NSColor.controlAccentColor.cgColor); ctx.setLineWidth(1.5)
            ctx.move(to: CGPoint(x: rx, y: vis.minY)); ctx.addLine(to: CGPoint(x: rx, y: vis.maxY)); ctx.strokePath()
        }
        ctx.restoreGState()
    }

    private func rangeRect(_ r: CellRange) -> NSRect {
        let co = coordinator
        let x1 = co.xForCol(r.start.column); let y1 = co.yForRow(r.start.row)
        let x2 = co.xForCol(r.end.column) + co.colW(r.end.column)
        let y2 = co.yForRow(r.end.row) + co.rowH(r.end.row)
        return NSRect(x: x1, y: y1, width: x2 - x1, height: y2 - y1)
    }

    private func drawHeaderGradient(ctx: CGContext, rect: NSRect, vertical: Bool) {
        let t = NSColor.white.withAlphaComponent(0.08).cgColor
        let b = NSColor.black.withAlphaComponent(0.04).cgColor
        guard let g = CGGradient(colorsSpace: CGColorSpaceCreateDeviceRGB(), colors: [t, b] as CFArray, locations: [0, 1]) else { return }
        ctx.saveGState(); ctx.clip(to: rect)
        if vertical {
            ctx.drawLinearGradient(g, start: CGPoint(x: rect.midX, y: rect.minY), end: CGPoint(x: rect.midX, y: rect.maxY), options: [])
        } else {
            ctx.drawLinearGradient(g, start: CGPoint(x: rect.minX, y: rect.midY), end: CGPoint(x: rect.maxX, y: rect.midY), options: [])
        }
        ctx.restoreGState()
    }

    // MARK: Mouse handling

    override func mouseDown(with event: NSEvent) {
        window?.makeFirstResponder(self)
        let loc = convert(event.locationInWindow, from: nil)
        let co = coordinator
        // Column resize in header
        if loc.y <= co.chh, let col = co.colBorderHit(at: loc.x) {
            co.isResizingColumn = true; co.resizeColumnIndex = col
            co.resizeStartX = loc.x; co.resizeStartWidth = co.colW(col); return
        }
        guard loc.x > co.rhw, loc.y > co.chh else { return }
        let addr = CellAddress(column: co.colAt(loc.x), row: co.rowAt(loc.y))
        if event.clickCount == 2 { co.viewModel.selectCell(addr); co.beginEditing(in: self); return }
        if co.viewModel.isEditing { co.commitAndRemoveEditor() }
        if event.modifierFlags.contains(.shift) { co.viewModel.selection.extendSelection(to: addr); co.viewModel.version += 1 }
        else { co.viewModel.selectCell(addr) }
        needsDisplay = true
    }

    override func mouseDragged(with event: NSEvent) {
        let loc = convert(event.locationInWindow, from: nil); let co = coordinator
        if co.isResizingColumn {
            let nw = max(20, co.resizeStartWidth + loc.x - co.resizeStartX)
            co.viewModel.activeSheet.setColumnWidth(Double(nw), for: co.resizeColumnIndex)
            co.viewModel.version += 1
            let tw = GridForgeSpacing.rowHeaderWidth + CGFloat(co.sheet.totalWidth(columns: co.nCols))
            if frame.width != tw { frame = NSRect(x: 0, y: 0, width: tw, height: frame.height) }
            needsDisplay = true; return
        }
        let addr = CellAddress(column: co.colAt(loc.x), row: co.rowAt(loc.y))
        co.viewModel.selection.extendSelection(to: addr); co.viewModel.version += 1; needsDisplay = true
    }

    override func mouseUp(with event: NSEvent) {
        if coordinator.isResizingColumn { coordinator.isResizingColumn = false; needsDisplay = true }
    }

    override func mouseMoved(with event: NSEvent) {
        let loc = convert(event.locationInWindow, from: nil); let co = coordinator
        if loc.y <= co.chh && loc.x > co.rhw {
            NSCursor.resizeLeftRight.set(co.colBorderHit(at: loc.x) != nil)
        } else { NSCursor.arrow.set() }
        if loc.x > co.rhw && loc.y > co.chh {
            let h = CellAddress(column: co.colAt(loc.x), row: co.rowAt(loc.y))
            if co.hoverCell != h { co.hoverCell = h; needsDisplay = true }
        } else if co.hoverCell != nil { co.hoverCell = nil; needsDisplay = true }
    }

    override func mouseExited(with event: NSEvent) {
        if coordinator.hoverCell != nil { coordinator.hoverCell = nil; needsDisplay = true }
        NSCursor.arrow.set()
    }

    // MARK: Right-click

    override func rightMouseDown(with event: NSEvent) {
        let loc = convert(event.locationInWindow, from: nil); let co = coordinator
        if loc.x > co.rhw && loc.y > co.chh {
            let addr = CellAddress(column: co.colAt(loc.x), row: co.rowAt(loc.y))
            if !co.viewModel.selection.selectedRange.contains(addr) { co.viewModel.selectCell(addr); needsDisplay = true }
        }
        NSMenu.popUpContextMenu(co.buildContextMenu(), with: event, for: self)
    }

    @objc func ctxCut(_ s: Any?) { coordinator.viewModel.copy(); coordinator.viewModel.deleteSelectedCells(); needsDisplay = true }
    @objc func ctxCopy(_ s: Any?) { coordinator.viewModel.copy() }
    @objc func ctxPaste(_ s: Any?) { coordinator.viewModel.paste(); needsDisplay = true }
    @objc func ctxInsRowAbove(_ s: Any?) { coordinator.viewModel.insertRow(); needsDisplay = true }
    @objc func ctxInsRowBelow(_ s: Any?) {
        let vm = coordinator.viewModel
        vm.selectCell(CellAddress(column: vm.activeCell.column, row: vm.activeCell.row + 1))
        vm.insertRow(); needsDisplay = true
    }
    @objc func ctxInsColLeft(_ s: Any?) { coordinator.viewModel.insertColumn(); needsDisplay = true }
    @objc func ctxInsColRight(_ s: Any?) {
        let vm = coordinator.viewModel
        vm.selectCell(CellAddress(column: vm.activeCell.column + 1, row: vm.activeCell.row))
        vm.insertColumn(); needsDisplay = true
    }
    @objc func ctxDelRow(_ s: Any?) { coordinator.viewModel.deleteRow(); needsDisplay = true }
    @objc func ctxDelCol(_ s: Any?) { coordinator.viewModel.deleteColumn(); needsDisplay = true }
    @objc func ctxClear(_ s: Any?) { coordinator.viewModel.deleteSelectedCells(); needsDisplay = true }

    // MARK: Keyboard handling

    override func keyDown(with event: NSEvent) {
        let vm = coordinator.viewModel; let co = coordinator
        let shift = event.modifierFlags.contains(.shift); let cmd = event.modifierFlags.contains(.command)

        if cmd {
            switch event.charactersIgnoringModifiers {
            case "c": vm.copy(); return
            case "x": vm.copy(); vm.deleteSelectedCells(); needsDisplay = true; return
            case "v": vm.paste(); needsDisplay = true; return
            case "z": if shift { vm.redo() } else { vm.undo() }; needsDisplay = true; return
            case "a": vm.selectAll(); needsDisplay = true; return
            case "b": vm.toggleBold(); needsDisplay = true; return
            case "i": vm.toggleItalic(); needsDisplay = true; return
            default: break
            }
            if event.keyCode == 115 { // Cmd+Home → A1
                vm.selectCell(CellAddress(column: 0, row: 0)); co.scrollToActiveCell(); needsDisplay = true; return
            }
        }

        switch event.keyCode {
        case 126: vm.moveSelection(direction: .up, extend: shift); co.scrollToActiveCell(); needsDisplay = true; return
        case 125: vm.moveSelection(direction: .down, extend: shift); co.scrollToActiveCell(); needsDisplay = true; return
        case 123: vm.moveSelection(direction: .left, extend: shift); co.scrollToActiveCell(); needsDisplay = true; return
        case 124: vm.moveSelection(direction: .right, extend: shift); co.scrollToActiveCell(); needsDisplay = true; return
        case 36: // Return
            if vm.isEditing { co.commitAndRemoveEditor(); vm.moveSelection(direction: .down, extend: false) }
            else { co.beginEditing(in: self) }
            needsDisplay = true; return
        case 48: // Tab
            if vm.isEditing { co.commitAndRemoveEditor() }
            vm.moveSelection(direction: .right, extend: false); co.scrollToActiveCell(); needsDisplay = true; return
        case 53: if vm.isEditing { co.cancelAndRemoveEditor(); needsDisplay = true }; return
        case 51, 117: if !vm.isEditing { vm.deleteSelectedCells(); needsDisplay = true }; return
        case 120: // F2
            if !vm.isEditing { co.beginEditing(in: self); needsDisplay = true }; return
        case 116: // Page Up
            pageMove(direction: .up, shift: shift); return
        case 121: // Page Down
            pageMove(direction: .down, shift: shift); return
        case 115: // Home (no cmd)
            vm.selectCell(CellAddress(column: 0, row: vm.activeCell.row)); co.scrollToActiveCell(); needsDisplay = true; return
        case 119: // End
            let lastCol = vm.activeSheet.usedRange?.end.column ?? 0
            vm.selectCell(CellAddress(column: lastCol, row: vm.activeCell.row)); co.scrollToActiveCell(); needsDisplay = true; return
        default: break
        }

        if let chars = event.characters, !chars.isEmpty, !cmd, let ch = chars.first, ch.isPrintable {
            co.beginEditing(in: self, withText: String(ch)); return
        }
        super.keyDown(with: event)
    }

    private func pageMove(direction: Direction, shift: Bool) {
        guard let sv = coordinator.scrollView else { return }
        let vm = coordinator.viewModel
        let rows = max(1, Int((sv.documentVisibleRect.height - coordinator.chh) / CGFloat(Worksheet.defaultRowHeight)))
        for _ in 0..<rows { vm.moveSelection(direction: direction, extend: shift) }
        coordinator.scrollToActiveCell(); needsDisplay = true
    }
}

// MARK: - NSCursor convenience

private extension NSCursor {
    func set(_ condition: Bool) { if condition { self.set() } else { NSCursor.arrow.set() } }
}

// MARK: - Character helpers

private extension Character {
    var isPrintable: Bool {
        guard let s = unicodeScalars.first else { return false }
        return s.value >= 32 && s.value < 127 || s.properties.isAlphabetic || s.properties.isEmoji
    }
}
