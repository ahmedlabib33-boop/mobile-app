import SwiftUI
import WebKit

struct ContentView: View {
    private let config = AppConfig.load()
    @State private var isLoading = true
    @State private var errorMessage: String?
    @State private var webView = WKWebView(frame: .zero, configuration: WebViewFactory.configuration())

    var body: some View {
        NavigationStack {
            ZStack {
                Color(.systemGroupedBackground).ignoresSafeArea()

                if let url = config.configuredURL {
                    WebView(
                        webView: webView,
                        url: url,
                        isLoading: $isLoading,
                        errorMessage: $errorMessage
                    )
                    .ignoresSafeArea(edges: .bottom)

                    if isLoading {
                        LoadingPanel()
                    }
                } else {
                    StatusPanel(
                        title: "Streamlit URL not configured",
                        message: "Edit Config/mobile_config.json and replace the placeholder with the deployed HTTPS Streamlit URL."
                    )
                }

                if let errorMessage {
                    StatusPanel(
                        title: "Dashboard unavailable",
                        message: errorMessage,
                        actionTitle: "Try Again",
                        action: reload
                    )
                }
            }
            .navigationTitle("Project Intelligence Hub")
            .navigationBarTitleDisplayMode(.inline)
            .toolbar {
                ToolbarItem(placement: .topBarTrailing) {
                    Button(action: reload) {
                        Image(systemName: "arrow.clockwise")
                    }
                    .accessibilityLabel("Refresh")
                }
            }
        }
    }

    private func reload() {
        errorMessage = nil
        isLoading = true
        webView.reload()
    }
}

struct LoadingPanel: View {
    var body: some View {
        VStack(spacing: 14) {
            ProgressView()
            Text("Loading Project Intelligence Hub")
                .font(.headline)
                .multilineTextAlignment(.center)
        }
        .padding(24)
        .background(.regularMaterial)
        .clipShape(RoundedRectangle(cornerRadius: 14, style: .continuous))
        .shadow(radius: 16)
    }
}

struct StatusPanel: View {
    let title: String
    let message: String
    var actionTitle: String?
    var action: (() -> Void)?

    var body: some View {
        VStack(spacing: 14) {
            Image(systemName: "wifi.exclamationmark")
                .font(.system(size: 44, weight: .semibold))
                .foregroundStyle(.teal)
            Text(title)
                .font(.title3.weight(.bold))
                .multilineTextAlignment(.center)
            Text(message)
                .font(.body)
                .foregroundStyle(.secondary)
                .multilineTextAlignment(.center)
            if let actionTitle, let action {
                Button(actionTitle, action: action)
                    .buttonStyle(.borderedProminent)
                    .controlSize(.large)
            }
        }
        .padding(24)
        .frame(maxWidth: 420)
        .background(Color(.systemBackground))
        .clipShape(RoundedRectangle(cornerRadius: 16, style: .continuous))
        .padding(24)
    }
}
