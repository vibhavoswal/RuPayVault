import Foundation

public struct WhatsAppFormatter {
    
    /// Formats vouchers for WhatsApp with bolding and monospace blocks
    public static func formatForWhatsApp(vouchers: [Voucher]) -> String {
        guard !vouchers.isEmpty else { return "" }
        let brand = vouchers.first?.brand ?? "Vouchers"
        let total = vouchers.reduce(0) { $0 + $1.amount }
        
        var text = "*\(brand) (\(vouchers.count) Vouchers — Rs. \(total.formatted()))*\n"
        text += "----------------------------------------\n"
        
        for (i, v) in vouchers.enumerated() {
            text += "\n\(i + 1). *Code:* `\(v.code)`"
            if let pin = v.pin, !pin.isEmpty {
                text += "\n   *PIN:* `\(pin)`"
            }
            text += "\n   *Valid:* \(v.dateOfIssue)\n"
        }
        
        text += "----------------------------------------\n"
        text += "*Total:* Rs. \(total.formatted())"
        return text
    }
    
    /// Clean plaintext for cashiers / billing operators
    public static func formatForPlaintext(vouchers: [Voucher]) -> String {
        guard !vouchers.isEmpty else { return "" }
        let brand = vouchers.first?.brand ?? "Vouchers"
        let total = vouchers.reduce(0) { $0 + $1.amount }
        
        var text = "\(brand) (\(vouchers.count) Total - Rs. \(total.formatted()))\n"
        text += "----------------------------------------\n"
        for (i, v) in vouchers.enumerated() {
            text += "\(i + 1). Code: \(v.code)"
            if let pin = v.pin, !pin.isEmpty {
                text += " | PIN: \(pin)"
            }
            text += "\n"
        }
        text += "----------------------------------------\n"
        text += "Total: Rs. \(total.formatted())"
        return text
    }
}
