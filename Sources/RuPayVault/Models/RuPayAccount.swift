import Foundation

public struct RuPayAccount: Identifiable, Codable, Sendable {
    public var id: String { username }
    public var username: String // e.g. "vibhavoswal9"
    public var lastRedeemedAt: Date
    public var totalRedemptions: Int
    
    public init(username: String, lastRedeemedAt: Date, totalRedemptions: Int = 0) {
        self.username = username
        self.lastRedeemedAt = lastRedeemedAt
        self.totalRedemptions = totalRedemptions
    }
    
    public var isCooldownReady: Bool {
        let elapsed = Date().timeIntervalSince(lastRedeemedAt)
        return elapsed >= (24 * 3600)
    }
    
    public var remainingCooldownFormatted: String {
        let remaining = (24 * 3600) - Date().timeIntervalSince(lastRedeemedAt)
        if remaining <= 0 { return "Ready" }
        let hours = Int(remaining) / 3600
        let minutes = (Int(remaining) % 3600) / 60
        return "\(hours)h \(minutes)m"
    }
}
