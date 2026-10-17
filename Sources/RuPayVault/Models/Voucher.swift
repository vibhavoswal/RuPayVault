import Foundation

public enum VoucherStatus: String, Codable, Sendable {
    case active = "ACTIVE"
    case pending = "PENDING"
    case used = "USED"
    case expired = "EXPIRED"
}

public struct Voucher: Identifiable, Codable, Sendable {
    public var id: UUID
    public var bookingId: String
    public var brand: String
    public var code: String
    public var pin: String?
    public var amount: Int
    public var dateOfIssue: String
    public var account: String
    public var cardNickname: String
    public var status: VoucherStatus
    public var redeemedDate: Date?
    
    public init(
        id: UUID = UUID(),
        bookingId: String,
        brand: String,
        code: String,
        pin: String? = nil,
        amount: Int,
        dateOfIssue: String,
        account: String,
        cardNickname: String,
        status: VoucherStatus = .active,
        redeemedDate: Date? = nil
    ) {
        self.id = id
        self.bookingId = bookingId
        self.brand = brand
        self.code = code
        self.pin = pin
        self.amount = amount
        self.dateOfIssue = dateOfIssue
        self.account = account
        self.cardNickname = cardNickname
        self.status = status
        self.redeemedDate = redeemedDate
    }
}

public struct UnassignedRedemption: Identifiable, Codable, Sendable {
    public var id: String { bookingId }
    public var bookingId: String
    public var brand: String
    public var amount: Int
    public var account: String
    public var dateOfIssue: String
    
    public init(bookingId: String, brand: String, amount: Int, account: String, dateOfIssue: String) {
        self.bookingId = bookingId
        self.brand = brand
        self.amount = amount
        self.account = account
        self.dateOfIssue = dateOfIssue
    }
}
