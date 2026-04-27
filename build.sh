#!/bin/bash
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
APP_NAME="GridForge"
BUILD_DIR="$PROJECT_DIR/.build"
APP_BUNDLE="$PROJECT_DIR/$APP_NAME.app"
DMG_PATH="$PROJECT_DIR/$APP_NAME.dmg"
ICON_PATH="$PROJECT_DIR/Assets/GridForge.icns"
MODULE_CACHE_DIR="$BUILD_DIR/module-cache"

# Code signing / notarization
DEVELOPER_ID="${GRIDFORGE_DEVELOPER_ID:-Developer ID Application: Simon-Pierre Boucher (3YM54G49SN)}"
TEAM_ID="3YM54G49SN"
APPLE_ID="spbou4@icloud.com"
BUNDLE_ID="com.gridforge.app"
XLSX_UTI="org.openxmlformats.spreadsheetml.sheet"

sign_identity_available() {
    security find-identity -v -p codesigning 2>/dev/null | grep -F "$DEVELOPER_ID" >/dev/null
}

sign_app() {
    if sign_identity_available; then
        echo "    Signing with Developer ID: $DEVELOPER_ID"
        codesign --deep --force --options runtime \
            --sign "$DEVELOPER_ID" \
            --timestamp \
            --identifier "$BUNDLE_ID" \
            "$APP_BUNDLE" 2>&1
    else
        echo "    Developer ID identity not found. Using ad-hoc signing for local testing."
        codesign --deep --force \
            --sign - \
            --identifier "$BUNDLE_ID" \
            "$APP_BUNDLE" 2>&1
    fi
}

sign_dmg() {
    if sign_identity_available; then
        echo "[Step 4] Signing DMG..."
        codesign --force --sign "$DEVELOPER_ID" --timestamp "$DMG_PATH" 2>&1
        echo "    DMG signed."
    else
        echo "[Step 4] Skipping DMG signing. Developer ID identity not found."
    fi
}

echo "=== GridForge Build System ==="
echo ""

mkdir -p "$MODULE_CACHE_DIR"
export CLANG_MODULE_CACHE_PATH="$MODULE_CACHE_DIR"

case "${1:-build}" in
    build)
        echo "[1/3] Building $APP_NAME (release)..."
        cd "$PROJECT_DIR"
        swift build -c release 2>&1

        echo "[2/3] Creating app bundle..."
        rm -rf "$APP_BUNDLE"
        mkdir -p "$APP_BUNDLE/Contents/MacOS"
        mkdir -p "$APP_BUNDLE/Contents/Resources"

        # Copy executable
        EXEC_PATH=$(swift build -c release --show-bin-path)/$APP_NAME
        cp "$EXEC_PATH" "$APP_BUNDLE/Contents/MacOS/$APP_NAME"

        # Copy icon
        if [ -f "$ICON_PATH" ]; then
            cp "$ICON_PATH" "$APP_BUNDLE/Contents/Resources/AppIcon.icns"
            echo "    Icon embedded."
        fi

        # Generate Info.plist
        cat > "$APP_BUNDLE/Contents/Info.plist" << 'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleDevelopmentRegion</key>
    <string>en</string>
    <key>CFBundleExecutable</key>
    <string>GridForge</string>
    <key>CFBundleIconFile</key>
    <string>AppIcon</string>
    <key>CFBundleIdentifier</key>
    <string>com.gridforge.app</string>
    <key>CFBundleInfoDictionaryVersion</key>
    <string>6.0</string>
    <key>CFBundleName</key>
    <string>GridForge</string>
    <key>CFBundleDisplayName</key>
    <string>GridForge</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0.0</string>
    <key>CFBundleVersion</key>
    <string>1</string>
    <key>LSMinimumSystemVersion</key>
    <string>14.0</string>
    <key>NSHighResolutionCapable</key>
    <true/>
    <key>NSSupportsAutomaticTermination</key>
    <true/>
    <key>NSSupportsSuddenTermination</key>
    <true/>
    <key>CFBundleDocumentTypes</key>
    <array>
        <dict>
            <key>CFBundleTypeName</key>
            <string>Excel Workbook</string>
            <key>CFBundleTypeExtensions</key>
            <array>
                <string>xlsx</string>
            </array>
            <key>CFBundleTypeRole</key>
            <string>Editor</string>
            <key>LSItemContentTypes</key>
            <array>
                <string>org.openxmlformats.spreadsheetml.sheet</string>
            </array>
            <key>LSHandlerRank</key>
            <string>Alternate</string>
        </dict>
    </array>
    <key>UTImportedTypeDeclarations</key>
    <array>
        <dict>
            <key>UTTypeIdentifier</key>
            <string>org.openxmlformats.spreadsheetml.sheet</string>
            <key>UTTypeDescription</key>
            <string>Excel Workbook</string>
            <key>UTTypeConformsTo</key>
            <array>
                <string>com.pkware.zip-archive</string>
                <string>public.data</string>
            </array>
            <key>UTTypeTagSpecification</key>
            <dict>
                <key>public.filename-extension</key>
                <array>
                    <string>xlsx</string>
                </array>
                <key>public.mime-type</key>
                <array>
                    <string>application/vnd.openxmlformats-officedocument.spreadsheetml.sheet</string>
                </array>
            </dict>
        </dict>
    </array>
</dict>
</plist>
PLIST

        echo "[3/3] App bundle ready."
        echo ""
        echo "=== Build complete ==="
        echo "App bundle: $APP_BUNDLE"
        echo ""
        echo "To run:       open $APP_BUNDLE"
        echo "To sign+dmg:  $0 dmg"
        ;;

    sign)
        echo "Signing $APP_NAME.app..."
        sign_app

        echo "Verifying signature..."
        codesign --verify --deep --strict --verbose=2 "$APP_BUNDLE" 2>&1
        echo ""
        echo "=== Signing complete ==="
        ;;

    dmg)
        echo "=== Building signed DMG ==="
        echo ""

        # Step 1: Build if needed
        if [ ! -d "$APP_BUNDLE" ]; then
            echo "[Step 1] Building app first..."
            "$0" build
            echo ""
        else
            echo "[Step 1] App bundle exists, skipping build."
        fi

        # Step 2: Sign the app
        echo "[Step 2] Code signing..."
        sign_app
        echo "    Signature verified:"
        codesign --verify --deep --strict "$APP_BUNDLE" 2>&1 && echo "    OK"

        # Step 3: Create DMG
        echo "[Step 3] Creating DMG..."
        rm -f "$DMG_PATH"

        # Create temp folder for DMG contents
        DMG_STAGING="$PROJECT_DIR/.dmg_staging"
        rm -rf "$DMG_STAGING"
        mkdir -p "$DMG_STAGING"
        cp -R "$APP_BUNDLE" "$DMG_STAGING/"
        ln -s /Applications "$DMG_STAGING/Applications"

        # Create DMG
        hdiutil create -volname "GridForge" \
            -srcfolder "$DMG_STAGING" \
            -ov -format UDZO \
            "$DMG_PATH" 2>&1

        rm -rf "$DMG_STAGING"
        echo "    DMG created: $DMG_PATH"

        # Step 4: Sign the DMG if a Developer ID certificate is installed.
        sign_dmg

        echo ""
        echo "=== DMG ready ==="
        echo "DMG: $DMG_PATH"
        echo ""
        echo "To notarize a Developer-ID-signed DMG:  $0 notarize"
        ;;

    notarize)
        echo "=== Notarizing GridForge ==="
        echo ""

        if [ ! -f "$DMG_PATH" ]; then
            echo "ERROR: DMG not found. Run '$0 dmg' first."
            exit 1
        fi

        # Step 1: Submit for notarization
        echo "[Step 1] Submitting to Apple notarization service..."
        echo "    Apple ID: $APPLE_ID"
        echo "    Team ID:  $TEAM_ID"
        echo ""

        xcrun notarytool submit "$DMG_PATH" \
            --apple-id "$APPLE_ID" \
            --team-id "$TEAM_ID" \
            --password "${GRIDFORGE_NOTARY_PASSWORD:?Set GRIDFORGE_NOTARY_PASSWORD to your app-specific password.}" \
            --wait 2>&1

        # Step 2: Staple the notarization ticket
        echo ""
        echo "[Step 2] Stapling notarization ticket to DMG..."
        xcrun stapler staple "$DMG_PATH" 2>&1

        # Step 3: Verify
        echo ""
        echo "[Step 3] Verifying notarization..."
        spctl --assess --type open --context context:primary-signature --verbose=2 "$DMG_PATH" 2>&1 || true
        xcrun stapler validate "$DMG_PATH" 2>&1

        echo ""
        echo "=== Notarization complete ==="
        echo "DMG ready for distribution: $DMG_PATH"
        echo ""
        echo "Users can now download and open GridForge without Gatekeeper warnings."
        ;;

    run)
        echo "Building and running $APP_NAME..."
        cd "$PROJECT_DIR"
        swift build 2>&1
        echo "Launching..."
        swift run $APP_NAME
        ;;

    debug)
        echo "Building $APP_NAME (debug)..."
        cd "$PROJECT_DIR"
        swift build 2>&1
        echo ""
        echo "Debug build complete."
        echo "Run with: swift run $APP_NAME"
        ;;

    test)
        echo "Running tests..."
        cd "$PROJECT_DIR"
        DEVELOPER_DIR=/Applications/Xcode.app/Contents/Developer swift test 2>&1
        ;;

    clean)
        echo "Cleaning build artifacts..."
        cd "$PROJECT_DIR"
        swift package clean
        rm -rf "$APP_BUNDLE" "$DMG_PATH" .dmg_staging
        echo "Clean complete."
        ;;

    *)
        echo "Usage: $0 {build|sign|dmg|notarize|run|debug|test|clean}"
        echo ""
        echo "  build     - Release build + app bundle with icon"
        echo "  sign      - Code sign the app bundle"
        echo "  dmg       - Build + sign + create DMG"
        echo "  notarize  - Submit DMG to Apple notarization + staple"
        echo "  run       - Debug build + run"
        echo "  debug     - Debug build only"
        echo "  test      - Run unit tests"
        echo "  clean     - Remove build artifacts"
        exit 1
        ;;
esac
