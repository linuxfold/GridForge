// swift-tools-version: 5.9
import PackageDescription

let package = Package(
    name: "GridForge",
    platforms: [.macOS(.v14)],
    dependencies: [
        .package(url: "https://github.com/weichsel/ZIPFoundation.git", from: "0.9.19")
    ],
    targets: [
        .target(
            name: "GridForgeCore",
            dependencies: ["ZIPFoundation"],
            path: "Sources/GridForgeCore"
        ),
        .executableTarget(
            name: "GridForge",
            dependencies: ["GridForgeCore"],
            path: "Sources/GridForge"
        ),
        .testTarget(
            name: "GridForgeTests",
            dependencies: ["GridForgeCore"],
            path: "Tests/GridForgeTests"
        )
    ]
)
