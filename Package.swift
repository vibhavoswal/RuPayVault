// swift-tools-version: 5.9
import PackageDescription

let package = Package(
    name: "RuPayVault",
    platforms: [
        .macOS(.v14),
        .iOS(.v17)
    ],
    products: [
        .executable(
            name: "RuPayVault",
            targets: ["RuPayVault"]
        )
    ],
    dependencies: [],
    targets: [
        .executableTarget(
            name: "RuPayVault",
            dependencies: [],
            path: "Sources/RuPayVault"
        )
    ]
)
