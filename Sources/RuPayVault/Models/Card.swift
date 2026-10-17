import Foundation

public enum CardStatus: String, Codable, Sendable {
    case active
    case blocked
}

public struct RuPayCard: Identifiable, Codable, Sendable {
    public var id: UUID
    public var bank: String
    public var cardType: String
    public var bin: String
    public var last4Digits: String
    public var nickname: String
    public var status: CardStatus
    public var issuedDate: Date?
    public var blockedDate: Date?
    public var notes: String?
    
    // Quota counts for current quarter (e.g. Q4 2026)
    public var claimedCount: Int
    public var totalQuota: Int
    
    public init(
        id: UUID = UUID(),
        bank: String,
        cardType: String,
        bin: String,
        last4Digits: String,
        nickname: String,
        status: CardStatus = .active,
        issuedDate: Date? = nil,
        blockedDate: Date? = nil,
        notes: String? = nil,
        claimedCount: Int = 0,
        totalQuota: Int = 6
    ) {
        self.id = id
        self.bank = bank
        self.cardType = cardType
        self.bin = bin
        self.last4Digits = last4Digits
        self.nickname = nickname
        self.status = status
        self.issuedDate = issuedDate
        self.blockedDate = blockedDate
        self.notes = notes
        self.claimedCount = claimedCount
        self.totalQuota = totalQuota
    }
}
