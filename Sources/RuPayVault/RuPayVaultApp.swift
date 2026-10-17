import SwiftUI

@main
struct RuPayVaultApp: App {
    @StateObject private var syncService = GoogleSheetSyncService.shared
    
    // Sample pre-loaded state matching user's bank accounts & cards
    @State private var cards: [RuPayCard] = [
        RuPayCard(bank: "PNB", cardType: "PNB Imperial RSDC", bin: "65249710", last4Digits: "9102", nickname: "PNB Imperial Primary", status: .active, claimedCount: 3, totalQuota: 6),
        RuPayCard(bank: "Bank of Baroda", cardType: "BoB Bhoomi RSDC", bin: "652505", last4Digits: "3314", nickname: "Vibhav BoB DC", status: .active, claimedCount: 2, totalQuota: 5),
        RuPayCard(bank: "Canara Bank", cardType: "Canara RSDC", bin: "652550", last4Digits: "7828", nickname: "Canara DC 7828", status: .active, claimedCount: 2, totalQuota: 5),
        RuPayCard(bank: "Jana SFB", cardType: "Jana SFB RSDC", bin: "652520", last4Digits: "7891", nickname: "Jana SFB DC", status: .active, claimedCount: 3, totalQuota: 5),
        RuPayCard(bank: "Canara Bank", cardType: "Canara Select CC", bin: "653020", last4Digits: "1109", nickname: "Canara CC Primary", status: .active, claimedCount: 1, totalQuota: 2),
        
        // Blocked cards (reverse chronological)
        RuPayCard(bank: "PNB", cardType: "PNB Imperial RSDC", bin: "65249710", last4Digits: "4821", nickname: "PNB Imperial (Old)", status: .blocked, blockedDate: Calendar.current.date(byAdding: .day, value: -3, to: Date()), notes: "Q3 quotas exhausted"),
        RuPayCard(bank: "Canara Bank", cardType: "Canara RSDC", bin: "652550", last4Digits: "1120", nickname: "Canara DC 1 (Old)", status: .blocked, blockedDate: Calendar.current.date(byAdding: .day, value: -19, to: Date()), notes: "Blocked via mobile banking")
    ]
    
    @State private var accounts: [RuPayAccount] = [
        RuPayAccount(username: "vibhavoswal9", lastRedeemedAt: Calendar.current.date(byAdding: .hour, value: -26, to: Date())!, totalRedemptions: 15),
        RuPayAccount(username: "karthikrathod2067", lastRedeemedAt: Calendar.current.date(byAdding: .hour, value: -19, to: Date())!, totalRedemptions: 8),
        RuPayAccount(username: "vibhavoswal2", lastRedeemedAt: Calendar.current.date(byAdding: .hour, value: -4, to: Date())!, totalRedemptions: 6),
        RuPayAccount(username: "nikhildsouza369", lastRedeemedAt: Calendar.current.date(byAdding: .hour, value: -25, to: Date())!, totalRedemptions: 7),
        RuPayAccount(username: "arunaoswal3", lastRedeemedAt: Calendar.current.date(byAdding: .hour, value: -8, to: Date())!, totalRedemptions: 12)
    ]
    
    @State private var unassignedRedemptions: [UnassignedRedemption] = [
        UnassignedRedemption(bookingId: "RUPS-76869357", brand: "Uber", amount: 100, account: "vibhavoswal9", dateOfIssue: "16-Oct-2026"),
        UnassignedRedemption(bookingId: "RUPS-76869358", brand: "Lakme", amount: 1500, account: "karthikrathod2067", dateOfIssue: "16-Oct-2026")
    ]
    
    @State private var vouchers: [Voucher] = [
        Voucher(bookingId: "RUPS-7680011", brand: "Lakme", code: "1001460114752185", pin: "112972", amount: 1500, dateOfIssue: "14-Oct-2026", account: "karthikrathod2067", cardNickname: "PNB Imperial Primary"),
        Voucher(bookingId: "RUPS-7680012", brand: "Lakme", code: "1001460112216149", pin: "115119", amount: 1500, dateOfIssue: "14-Oct-2026", account: "karthikrathod2067", cardNickname: "PNB Imperial Primary"),
        Voucher(bookingId: "RUPS-7680021", brand: "Kalyan", code: "1004681040330454", pin: "108190", amount: 2000, dateOfIssue: "10-Oct-2026", account: "devendixit369", cardNickname: "PNB Imperial Primary"),
        Voucher(bookingId: "RUPS-7680031", brand: "Reliance", code: "GYFRR24673597621", pin: nil, amount: 500, dateOfIssue: "12-Oct-2026", account: "devendixit369", cardNickname: "PNB Imperial Primary"),
        Voucher(bookingId: "RUPS-7680041", brand: "Uber", code: "QDR22U3K4DTB5", pin: nil, amount: 100, dateOfIssue: "16-Oct-2026", account: "vibhavoswal9", cardNickname: "Canara DC 7828")
    ]
    
    var body: some Scene {
        WindowGroup {
            ContentView(
                cards: $cards,
                accounts: $accounts,
                unassignedRedemptions: $unassignedRedemptions,
                vouchers: $vouchers
            )
            .environmentObject(syncService)
        }
        #if os(macOS)
        .windowStyle(.hiddenTitleBar)
        .windowToolbarStyle(.unified)
        #endif
    }
}

// MARK: - Root Content View (Adaptive macOS & iOS)
struct ContentView: View {
    @Binding var cards: [RuPayCard]
    @Binding var accounts: [RuPayAccount]
    @Binding var unassignedRedemptions: [UnassignedRedemption]
    @Binding var vouchers: [Voucher]
    
    #if os(macOS)
    @State private var selectedSidebarTab: String? = "cards"
    
    var body: some View {
        NavigationSplitView {
            SidebarView(
                selectedTab: $selectedSidebarTab,
                activeCardsCount: cards.filter { $0.status == .active }.count,
                vouchersCount: vouchers.count,
                blockedCount: cards.filter { $0.status == .blocked }.count,
                unassignedCount: unassignedRedemptions.count
            )
        } detail: {
            switch selectedSidebarTab {
            case "unassigned":
                ReconciliationWizardView(cards: cards, unassigned: $unassignedRedemptions, vouchers: $vouchers)
            case "cards":
                CardsManagementView(cards: $cards)
            case "vault":
                VoucherVaultView(vouchers: $vouchers)
            case "cooldowns":
                CooldownsMonitorView(accounts: $accounts)
            default:
                CardsManagementView(cards: $cards)
            }
        }
        .frame(minWidth: 900, minHeight: 620)
    }
    #else
    // iPhone 16 Pro (Bottom Tab View + Dynamic Reachability)
    var body: some View {
        TabView {
            NavigationStack {
                VStack(spacing: 0) {
                    if !unassignedRedemptions.isEmpty {
                        NavigationLink {
                            ReconciliationWizardView(cards: cards, unassigned: $unassignedRedemptions, vouchers: $vouchers)
                        } label: {
                            HStack {
                                Image(systemName: "bolt.fill")
                                    .foregroundStyle(.yellow)
                                Text("\(unassignedRedemptions.count) Redemptions Waiting for Card")
                                    .font(.subheadline.weight(.semibold))
                                Spacer()
                                Image(systemName: "chevron.right")
                                    .font(.caption)
                                    .foregroundStyle(.secondary)
                            }
                            .padding(.horizontal, 16)
                            .padding(.vertical, 10)
                            .background(Color.yellow.opacity(0.12))
                        }
                    }
                    CardsManagementView(cards: $cards)
                }
                .navigationTitle("Cards")
            }
            .tabItem { Label("Cards", systemImage: "creditcard") }
            
            NavigationStack {
                VoucherVaultView(vouchers: $vouchers)
                    .navigationTitle("Vault")
            }
            .tabItem { Label("Vault", systemImage: "archivebox") }
            
            NavigationStack {
                CooldownsMonitorView(accounts: $accounts)
                    .navigationTitle("Cooldowns")
            }
            .tabItem { Label("Cooldowns", systemImage: "timer") }
        }
    }
    #endif
}

// MARK: - Reconciliation Wizard (Rapid 1-Tap Card Tagging)
struct ReconciliationWizardView: View {
    let cards: [RuPayCard]
    @Binding var unassigned: [UnassignedRedemption]
    @Binding var vouchers: [Voucher]
    
    @State private var batchSelectedCardId: UUID? = nil
    @State private var individualAssignments: [String: UUID] = [:]
    
    var activeCards: [RuPayCard] {
        cards.filter { $0.status == .active }
    }
    
    var body: some View {
        ScrollView {
            VStack(alignment: .leading, spacing: 16) {
                // Header Banner
                VStack(alignment: .leading, spacing: 4) {
                    Text("Reconcile Redemptions")
                        .font(.title2.weight(.bold))
                    Text("Emails received from rupay@truztee.com with no card number specified. Select the card used to update quotas.")
                        .font(.footnote)
                        .foregroundStyle(.secondary)
                }
                .padding(.bottom, 8)
                
                // Batch "Apply to All" Bar
                HStack(spacing: 12) {
                    Text("Bulk Assign:")
                        .font(.subheadline.weight(.medium))
                    
                    Picker("Card", selection: $batchSelectedCardId) {
                        Text("Select Card for All").tag(Optional<UUID>.none)
                        ForEach(activeCards) { card in
                            Text("\(card.bank) (\(card.last4Digits))").tag(Optional(card.id))
                        }
                    }
                    .pickerStyle(.menu)
                    
                    Button("Apply All") {
                        if let batchId = batchSelectedCardId {
                            for item in unassigned {
                                individualAssignments[item.bookingId] = batchId
                            }
                        }
                    }
                    .buttonStyle(.borderedProminent)
                    .disabled(batchSelectedCardId == nil)
                }
                .padding()
                .background(Color(.secondarySystemBackground))
                .clipShape(RoundedRectangle(cornerRadius: 12))
                
                // Individual Items List
                VStack(spacing: 8) {
                    ForEach(unassigned) { item in
                        HStack {
                            VStack(alignment: .leading, spacing: 2) {
                                HStack {
                                    Text(item.brand)
                                        .font(.headline)
                                    Text("Rs. \(item.amount)")
                                        .font(.subheadline.weight(.semibold))
                                        .foregroundStyle(.green)
                                }
                                Text("ID: \(item.bookingId) • Account: \(item.account)")
                                    .font(.caption)
                                    .foregroundStyle(.secondary)
                            }
                            Spacer()
                            
                            Picker("Card", selection: Binding(
                                get: { individualAssignments[item.bookingId] },
                                set: { individualAssignments[item.bookingId] = $0 }
                            )) {
                                Text("Pick Card").tag(Optional<UUID>.none)
                                ForEach(activeCards) { card in
                                    Text("\(card.bank) ···· \(card.last4Digits)").tag(Optional(card.id))
                                }
                            }
                            .pickerStyle(.menu)
                        }
                        .padding(12)
                        .background(Color(.secondarySystemBackground))
                        .clipShape(RoundedRectangle(cornerRadius: 10))
                    }
                }
                
                // Submit Button
                Button {
                    commitAssignments()
                } label: {
                    Text("Confirm & Sync to Google Sheets")
                        .font(.headline)
                        .frame(maxWidth: .infinity)
                        .padding(.vertical, 8)
                }
                .buttonStyle(.borderedProminent)
                .padding(.top, 12)
            }
            .padding()
        }
    }
    
    private func commitAssignments() {
        for item in unassigned {
            if let cardId = individualAssignments[item.bookingId],
               let card = cards.first(where: { $0.id == cardId }) {
                vouchers.append(
                    Voucher(
                        bookingId: item.bookingId,
                        brand: item.brand,
                        code: "PENDING_DELIVERY",
                        pin: nil,
                        amount: item.amount,
                        dateOfIssue: item.dateOfIssue,
                        account: item.account,
                        cardNickname: "\(card.bank) ···· \(card.last4Digits)",
                        status: .pending
                    )
                )
            }
        }
        unassigned.removeAll()
    }
}

// MARK: - Cards Management View (Active + Blocked)
struct CardsManagementView: View {
    @Binding var cards: [RuPayCard]
    @State private var isBlockedExpanded: Bool = false
    @State private var cardToBlock: RuPayCard? = nil
    @State private var blockDate: Date = Date()
    @State private var blockNotes: String = ""
    
    var activeCards: [RuPayCard] {
        cards.filter { $0.status == .active }
    }
    
    var blockedCards: [RuPayCard] {
        cards.filter { $0.status == .blocked }.sorted { ($0.blockedDate ?? Date.distantPast) > ($1.blockedDate ?? Date.distantPast) }
    }
    
    var body: some View {
        List {
            Section("Active Cards (\(activeCards.count))") {
                ForEach(activeCards) { card in
                    VStack(alignment: .leading, spacing: 8) {
                        HStack {
                            VStack(alignment: .leading, spacing: 2) {
                                Text(card.bank.uppercased())
                                    .font(.caption2.weight(.bold))
                                    .foregroundStyle(.blue)
                                Text(card.nickname)
                                    .font(.headline)
                                Text("BIN \(card.bin) • ···· \(card.last4Digits)")
                                    .font(.caption.monospaced())
                                    .foregroundStyle(.secondary)
                            }
                            Spacer()
                            Text("Active")
                                .font(.caption.weight(.semibold))
                                .foregroundStyle(.green)
                                .padding(.horizontal, 8)
                                .padding(.vertical, 3)
                                .background(Color.green.opacity(0.12))
                                .clipShape(Capsule())
                        }
                        
                        // Quota Progress Bar
                        ProgressView(value: Double(card.claimedCount), total: Double(card.totalQuota))
                            .tint(.blue)
                        
                        HStack {
                            Text("Q4 Quotas: \(card.claimedCount)/\(card.totalQuota) claimed")
                                .font(.caption)
                                .foregroundStyle(.secondary)
                            Spacer()
                            Button("Block Card", role: .destructive) {
                                cardToBlock = card
                                blockDate = Date()
                            }
                            .font(.caption.weight(.medium))
                        }
                    }
                    .padding(.vertical, 4)
                }
            }
            
            // Blocked Cards (Reverse Sorted by Block Date)
            Section(isExpanded: $isBlockedExpanded) {
                ForEach(blockedCards) { card in
                    HStack {
                        VStack(alignment: .leading, spacing: 2) {
                            Text(card.nickname)
                                .font(.subheadline.weight(.medium))
                                .strikethrough()
                                .foregroundStyle(.red)
                            if let bDate = card.blockedDate {
                                Text("Blocked on \(bDate.formatted(date: .abbreviated, time: .omitted)) • \(card.notes ?? "")")
                                    .font(.caption)
                                    .foregroundStyle(.secondary)
                            }
                        }
                        Spacer()
                        Text("Blocked")
                            .font(.caption2.weight(.bold))
                            .foregroundStyle(.red)
                            .padding(.horizontal, 6)
                            .padding(.vertical, 2)
                            .background(Color.red.opacity(0.1))
                            .clipShape(Capsule())
                    }
                }
            } header: {
                HStack {
                    Text("Blocked Cards Archive (\(blockedCards.count))")
                    Spacer()
                    Image(systemName: isBlockedExpanded ? "chevron.down" : "chevron.right")
                        .font(.caption)
                }
                .contentShape(Rectangle())
                .onTapGesture {
                    withAnimation { isBlockedExpanded.toggle() }
                }
            }
        }
        .sheet(item: $cardToBlock) { card in
            NavigationStack {
                Form {
                    Section("Card Details") {
                        Text(card.nickname).font(.headline)
                        Text("···· \(card.last4Digits)").font(.caption.monospaced())
                    }
                    Section("Date of Block") {
                        DatePicker("Block Date", selection: $blockDate, displayedComponents: .date)
                        TextField("Notes / Reason", text: $blockNotes)
                    }
                }
                .navigationTitle("Block Card")
                .toolbar {
                    ToolbarItem(placement: .cancellationAction) {
                        Button("Cancel") { cardToBlock = nil }
                    }
                    ToolbarItem(placement: .confirmationAction) {
                        Button("Confirm Block", role: .destructive) {
                            if let idx = cards.firstIndex(where: { $0.id == card.id }) {
                                cards[idx].status = .blocked
                                cards[idx].blockedDate = blockDate
                                cards[idx].notes = blockNotes
                            }
                            cardToBlock = nil
                        }
                    }
                }
            }
        }
    }
}

// MARK: - Voucher Vault (Multi-Select & WhatsApp)
struct VoucherVaultView: View {
    @Binding var vouchers: [Voucher]
    @State private var selectedIds = Set<UUID>()
    @State private var isSelectionMode: Bool = false
    
    var selectedVouchers: [Voucher] {
        vouchers.filter { selectedIds.contains($0.id) }
    }
    
    var totalSelectedSum: Int {
        selectedVouchers.reduce(0) { $0 + $1.amount }
    }
    
    var body: some View {
        VStack(spacing: 0) {
            List(selection: isSelectionMode ? $selectedIds : .constant([])) {
                ForEach(vouchers) { voucher in
                    HStack {
                        VStack(alignment: .leading, spacing: 3) {
                            HStack {
                                Text(voucher.brand).font(.headline)
                                Text("Rs. \(voucher.amount)").font(.subheadline.weight(.semibold)).foregroundStyle(.green)
                            }
                            Text("Code: \(voucher.code)").font(.system(.subheadline, design: .monospaced))
                            if let pin = voucher.pin {
                                Text("PIN: \(pin)").font(.caption.monospaced()).foregroundStyle(.secondary)
                            }
                            Text("Account: \(voucher.account) • \(voucher.cardNickname)").font(.caption2).foregroundStyle(.secondary)
                        }
                        Spacer()
                        VStack(alignment: .trailing, spacing: 4) {
                            Text(voucher.status.rawValue)
                                .font(.caption2.weight(.bold))
                                .foregroundStyle(voucher.status == .active ? .green : .secondary)
                            Text(voucher.dateOfIssue)
                                .font(.caption2)
                                .foregroundStyle(.secondary)
                        }
                    }
                }
            }
            .environment(\.editMode, .constant(isSelectionMode ? .active : .inactive))
            
            // Bottom Multi-Select Actions (WhatsApp & Plaintext)
            if isSelectionMode && !selectedIds.isEmpty {
                HStack(spacing: 12) {
                    VStack(alignment: .leading) {
                        Text("\(selectedIds.count) Selected")
                            .font(.subheadline.weight(.semibold))
                        Text("Rs. \(totalSelectedSum.formatted()) Total")
                            .font(.caption)
                            .foregroundStyle(.secondary)
                    }
                    Spacer()
                    
                    Button {
                        #if os(macOS)
                        NSPasteboard.general.clearContents()
                        NSPasteboard.general.setString(WhatsAppFormatter.formatForPlaintext(vouchers: selectedVouchers), forType: .string)
                        #else
                        UIPasteboard.general.string = WhatsAppFormatter.formatForPlaintext(vouchers: selectedVouchers)
                        #endif
                    } label: {
                        Label("Plain", systemImage: "doc.on.doc")
                    }
                    .buttonStyle(.bordered)
                    
                    ShareLink(item: WhatsAppFormatter.formatForWhatsApp(vouchers: selectedVouchers)) {
                        Label("WhatsApp", systemImage: "square.and.arrow.up")
                    }
                    .buttonStyle(.borderedProminent)
                    .tint(.green)
                }
                .padding()
                .background(.ultraThinMaterial)
            }
        }
        .toolbar {
            ToolbarItem(placement: .primaryAction) {
                Button(isSelectionMode ? "Done" : "Select") {
                    withAnimation {
                        isSelectionMode.toggle()
                        if !isSelectionMode { selectedIds.removeAll() }
                    }
                }
            }
        }
    }
}

// MARK: - Cooldowns Monitor View
struct CooldownsMonitorView: View {
    @Binding var accounts: [RuPayAccount]
    
    var body: some View {
        List(accounts) { account in
            HStack {
                VStack(alignment: .leading, spacing: 2) {
                    Text(account.username)
                        .font(.headline.monospaced())
                    Text("Redemptions: \(account.totalRedemptions)")
                        .font(.caption)
                        .foregroundStyle(.secondary)
                }
                Spacer()
                Text(account.remainingCooldownFormatted)
                    .font(.subheadline.weight(.semibold))
                    .foregroundStyle(account.isCooldownReady ? .green : .orange)
                    .padding(.horizontal, 10)
                    .padding(.vertical, 4)
                    .background(account.isCooldownReady ? Color.green.opacity(0.12) : Color.orange.opacity(0.12))
                    .clipShape(Capsule())
            }
        }
    }
}

#if os(macOS)
struct SidebarView: View {
    @Binding var selectedTab: String?
    let activeCardsCount: Int
    let vouchersCount: Int
    let blockedCount: Int
    let unassignedCount: Int
    
    var body: some View {
        List(selection: $selectedTab) {
            if unassignedCount > 0 {
                NavigationLink(value: "unassigned") {
                    Label("Reconcile (\(unassignedCount))", systemImage: "bolt.fill")
                        .foregroundStyle(.yellow)
                }
            }
            NavigationLink(value: "cards") {
                Label("Cards (\(activeCardsCount))", systemImage: "creditcard")
            }
            NavigationLink(value: "vault") {
                Label("Vault (\(vouchersCount))", systemImage: "archivebox")
            }
            NavigationLink(value: "cooldowns") {
                Label("24h Cooldowns", systemImage: "timer")
            }
        }
        .listStyle(.sidebar)
    }
}
#endif
