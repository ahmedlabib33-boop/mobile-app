import SwiftUI
import WebKit

enum WebViewFactory {
    static func configuration() -> WKWebViewConfiguration {
        let preferences = WKWebpagePreferences()
        preferences.allowsContentJavaScript = true

        let configuration = WKWebViewConfiguration()
        configuration.defaultWebpagePreferences = preferences
        configuration.websiteDataStore = .default()
        configuration.allowsInlineMediaPlayback = true
        return configuration
    }
}

struct WebView: UIViewRepresentable {
    let webView: WKWebView
    let url: URL
    @Binding var isLoading: Bool
    @Binding var errorMessage: String?

    func makeCoordinator() -> Coordinator {
        Coordinator(self)
    }

    func makeUIView(context: Context) -> WKWebView {
        webView.navigationDelegate = context.coordinator
        webView.uiDelegate = context.coordinator
        webView.allowsBackForwardNavigationGestures = true
        webView.scrollView.contentInsetAdjustmentBehavior = .automatic
        webView.load(URLRequest(url: url))
        return webView
    }

    func updateUIView(_ uiView: WKWebView, context: Context) {}

    final class Coordinator: NSObject, WKNavigationDelegate, WKUIDelegate {
        private let parent: WebView

        init(_ parent: WebView) {
            self.parent = parent
        }

        func webView(_ webView: WKWebView, didStartProvisionalNavigation navigation: WKNavigation!) {
            parent.isLoading = true
            parent.errorMessage = nil
        }

        func webView(_ webView: WKWebView, didFinish navigation: WKNavigation!) {
            parent.isLoading = false
        }

        func webView(_ webView: WKWebView, didFail navigation: WKNavigation!, withError error: Error) {
            parent.isLoading = false
            parent.errorMessage = "Check the internet connection and confirm the deployed dashboard URL is reachable."
        }

        func webView(_ webView: WKWebView, didFailProvisionalNavigation navigation: WKNavigation!, withError error: Error) {
            parent.isLoading = false
            parent.errorMessage = "Check the internet connection and confirm the deployed dashboard URL is reachable."
        }

        func webView(
            _ webView: WKWebView,
            decidePolicyFor navigationAction: WKNavigationAction,
            decisionHandler: @escaping (WKNavigationActionPolicy) -> Void
        ) {
            guard let targetURL = navigationAction.request.url else {
                decisionHandler(.cancel)
                return
            }

            let configuredHost = parent.url.host?.lowercased()
            let targetHost = targetURL.host?.lowercased()
            let isSameHost = configuredHost == targetHost
            let isExport = looksLikeExport(targetURL)

            if !isSameHost || isExport {
                UIApplication.shared.open(targetURL)
                decisionHandler(.cancel)
                return
            }

            decisionHandler(.allow)
        }

        private func looksLikeExport(_ url: URL) -> Bool {
            let path = url.path.lowercased()
            return path.hasSuffix(".pdf")
                || path.hasSuffix(".xlsx")
                || path.hasSuffix(".xls")
                || path.hasSuffix(".csv")
                || path.hasSuffix(".docx")
                || path.hasSuffix(".pptx")
                || path.hasSuffix(".zip")
                || path.contains("download")
        }
    }
}
