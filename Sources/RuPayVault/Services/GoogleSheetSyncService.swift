import Foundation
import Combine

@MainActor
public final class GoogleSheetSyncService: ObservableObject {
    public static let shared = GoogleSheetSyncService()
    
    // User live deployed Web App URL
    @Published public var webAppUrl: String = "https://script.google.com/macros/s/AKfycbxhdhzl1DpHnvGXYV0KaENylb6hUw-s_hzo39nsSf37dxwnSnuh-OPrWLlvSJvfvwpwZw/exec"
    @Published public var isSyncing: Bool = false
    @Published public var lastSyncDate: Date? = nil
    
    private init() {}
    
    /// Fetches all unassigned redemptions, cooldowns, and vouchers from Google Sheets
    public func fetchLatestVaultData() async throws -> (unassigned: [UnassignedRedemption], accounts: [RuPayAccount], vouchers: [Voucher]) {
        guard let url = URL(string: webAppUrl) else {
            return ([], [], [])
        }
        
        isSyncing = true
        defer { isSyncing = false }
        
        let (data, _) = try await URLSession.shared.data(from: url)
        
        struct ApiResponse: Codable {
            struct RawUnassigned: Codable {
                let bookingId: String
                let brand: String
                let amount: Int
                let account: String
                let dateOfIssue: String
            }
            struct RawCooldown: Codable {
                let account: String
                let lastRedeemed: String
            }
            struct RawVoucher: Codable {
                let brand: String
                let bookingId: String
                let code: String
                let pin: String?
                let amount: Int
                let dateOfIssue: String
                let account: String
                let card: String
                let status: String
            }
            
            let unassigned: [RawUnassigned]
            let cooldowns: [RawCooldown]
            let vouchers: [RawVoucher]
        }
        
        let decoded = try JSONDecoder().decode(ApiResponse.self, from: data)
        
        let unassigned = decoded.unassigned.map {
            UnassignedRedemption(bookingId: $0.bookingId, brand: $0.brand, amount: $0.amount, account: $0.account, dateOfIssue: $0.dateOfIssue)
        }
        
        let accounts = decoded.cooldowns.map {
            RuPayAccount(username: $0.account, lastRedeemedAt: ISO8601DateFormatter().date(from: $0.lastRedeemed) ?? Date())
        }
        
        let vouchers = decoded.vouchers.map {
            Voucher(
                bookingId: $0.bookingId,
                brand: $0.brand,
                code: $0.code,
                pin: $0.pin,
                amount: $0.amount,
                dateOfIssue: $0.dateOfIssue,
                account: $0.account,
                cardNickname: $0.card,
                status: VoucherStatus(rawValue: $0.status) ?? .active
            )
        }
        
        lastSyncDate = Date()
        return (unassigned, accounts, vouchers)
    }
    
    /// Assigns cards to unassigned batch redemptions in Google Sheets
    public func assignCardsToBatch(assignments: [(bookingId: String, card: String)]) async throws {
        guard let url = URL(string: webAppUrl) else { return }
        
        var request = URLRequest(url: url)
        request.httpMethod = "POST"
        request.setValue("application/json", forHTTPHeaderField: "Content-Type")
        
        let payload: [String: Any] = [
            "action": "ASSIGN_CARDS",
            "assignments": assignments.map { ["bookingId": $0.bookingId, "card": $0.card] }
        ]
        
        request.httpBody = try JSONSerialization.data(withJSONObject: payload)
        _ = try await URLSession.shared.data(for: request)
    }
    
    /// Marks voucher as used in Google Sheets
    public func markVoucherAsUsed(brand: String, bookingId: String, code: String) async throws {
        guard let url = URL(string: webAppUrl) else { return }
        
        var request = URLRequest(url: url)
        request.httpMethod = "POST"
        request.setValue("application/json", forHTTPHeaderField: "Content-Type")
        
        let payload: [String: Any] = [
            "action": "MARK_REDEEMED",
            "brand": brand,
            "bookingId": bookingId,
            "code": code
        ]
        
        request.httpBody = try JSONSerialization.data(withJSONObject: payload)
        _ = try await URLSession.shared.data(for: request)
    }
}
