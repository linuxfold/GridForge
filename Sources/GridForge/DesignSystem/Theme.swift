import SwiftUI
import AppKit

// MARK: - Adaptive Colors (Light + Dark mode)

enum GridForgeColors {
    // Grid
    static let gridLine = Color(nsColor: .separatorColor)
    static let headerBackground = Color(nsColor: .unemphasizedSelectedContentBackgroundColor)
    static let headerText = Color(nsColor: .secondaryLabelColor)
    static let activeHeaderText = Color(nsColor: .labelColor)
    static let cellBackground = Color(nsColor: .textBackgroundColor)
    static let cellText = Color(nsColor: .labelColor)

    // Selection
    static let selectedCellBorder = Color.accentColor
    static let selectedRangeFill = Color.accentColor.opacity(0.12)
    static let activeHeaderHighlight = Color.accentColor.opacity(0.18)

    // Chrome
    static let toolbarBackground = Color(nsColor: .windowBackgroundColor)
    static let formulaBarBackground = Color(nsColor: .controlBackgroundColor)
    static let formulaBarBorder = Color(nsColor: .separatorColor)
    static let inspectorBackground = Color(nsColor: .controlBackgroundColor)

    // Sheet Tabs
    static let sheetTabActive = Color.accentColor
    static let sheetTabActiveText = Color.white
    static let sheetTabInactive = Color(nsColor: .controlBackgroundColor)
    static let sheetTabInactiveText = Color(nsColor: .secondaryLabelColor)
    static let sheetTabHover = Color(nsColor: .unemphasizedSelectedContentBackgroundColor)

    // Status Bar
    static let statusBarBackground = Color(nsColor: .windowBackgroundColor)
    static let statusBarText = Color(nsColor: .secondaryLabelColor)

    // Semantic
    static let errorText = Color(nsColor: .systemRed)
    static let formulaText = Color(nsColor: .systemBlue)
    static let linkText = Color.accentColor
    static let disabledText = Color(nsColor: .tertiaryLabelColor)
    static let divider = Color(nsColor: .separatorColor)
}

// MARK: - NSColor equivalents (for AppKit / Core Graphics drawing in GridView)

enum GridForgeNSColors {
    // Grid
    static let gridLine = NSColor.separatorColor
    static let gridLineLight = NSColor.separatorColor.withAlphaComponent(0.4)
    static let headerBackground = NSColor.unemphasizedSelectedContentBackgroundColor
    static let headerBorder = NSColor.separatorColor
    static let headerText = NSColor.secondaryLabelColor
    static let activeHeaderText = NSColor.labelColor
    static let cellBackground = NSColor.textBackgroundColor
    static let cellText = NSColor.labelColor
    static let cornerBackground = NSColor.unemphasizedSelectedContentBackgroundColor

    // Selection
    static let selectedCellBorder = NSColor.controlAccentColor
    static let selectedRangeFill = NSColor.controlAccentColor.withAlphaComponent(0.12)
    static let selectedRangeBorder = NSColor.controlAccentColor.withAlphaComponent(0.35)
    static let activeHeaderHighlight = NSColor.controlAccentColor.withAlphaComponent(0.18)

    // Cell editor
    static let editorBackground = NSColor.textBackgroundColor
    static let editorBorder = NSColor.controlAccentColor
    static let editorShadow = NSColor.shadowColor.withAlphaComponent(0.12)

    // Hover
    static let cellHover = NSColor.labelColor.withAlphaComponent(0.04)
    static let headerHover = NSColor.labelColor.withAlphaComponent(0.06)

    // Error
    static let errorText = NSColor.systemRed
}

// MARK: - Typography

enum GridForgeTypography {
    // SwiftUI fonts
    static let cellFont = Font.system(size: 13)
    static let cellFontBold = Font.system(size: 13, weight: .semibold)
    static let headerFont = Font.system(size: 11, weight: .medium)
    static let formulaBarFont = Font.system(size: 13, design: .monospaced)
    static let formulaBarLabel = Font.system(size: 12, weight: .medium, design: .monospaced)
    static let toolbarFont = Font.system(size: 12)
    static let sheetTabFont = Font.system(size: 12, weight: .medium)
    static let inspectorHeading = Font.system(size: 11, weight: .semibold)
    static let inspectorLabel = Font.system(size: 11)
    static let inspectorValue = Font.system(size: 11, design: .monospaced)
    static let statusBarFont = Font.system(size: 11)
    static let nameBoxFont = Font.system(size: 12, weight: .medium, design: .monospaced)
}

enum GridForgeNSFonts {
    static let cellFont = NSFont.systemFont(ofSize: 13)
    static let cellFontBold = NSFont.systemFont(ofSize: 13, weight: .semibold)
    static let cellFontItalic: NSFont = {
        NSFontManager.shared.convert(cellFont, toHaveTrait: .italicFontMask)
    }()
    static let cellFontBoldItalic: NSFont = {
        NSFontManager.shared.convert(cellFontBold, toHaveTrait: .italicFontMask)
    }()
    static let headerFont = NSFont.systemFont(ofSize: 11, weight: .medium)
    static let editorFont = NSFont.systemFont(ofSize: 13)
}

// MARK: - Spacing & Layout

enum GridForgeSpacing {
    static let rowHeaderWidth: CGFloat = 52
    static let columnHeaderHeight: CGFloat = 26
    static let defaultCellWidth: CGFloat = 100
    static let defaultCellHeight: CGFloat = 24
    static let formulaBarHeight: CGFloat = 34
    static let sheetTabBarHeight: CGFloat = 30
    static let statusBarHeight: CGFloat = 24
    static let cellPaddingH: CGFloat = 6
    static let cellPaddingV: CGFloat = 3
    static let inspectorWidth: CGFloat = 250
    static let sheetTabMinWidth: CGFloat = 60
    static let sheetTabMaxWidth: CGFloat = 160
    static let sheetTabHeight: CGFloat = 24
    static let gridCornerRadius: CGFloat = 0
}

// MARK: - Animation Timing

enum GridForgeAnimation {
    static let quick: Animation = .easeOut(duration: 0.12)
    static let standard: Animation = .easeInOut(duration: 0.2)
    static let smooth: Animation = .easeInOut(duration: 0.3)
    static let spring: Animation = .interpolatingSpring(stiffness: 300, damping: 25)
}
