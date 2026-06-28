import Foundation

struct AppConfig: Decodable {
    let appName: String
    let streamlitUrl: String

    enum CodingKeys: String, CodingKey {
        case appName = "app_name"
        case streamlitUrl = "streamlit_url"
    }

    static let placeholderUrl = "PUT_DEPLOYED_STREAMLIT_URL_HERE"

    static func load() -> AppConfig {
        guard
            let url = Bundle.main.url(forResource: "mobile_config", withExtension: "json"),
            let data = try? Data(contentsOf: url),
            let config = try? JSONDecoder().decode(AppConfig.self, from: data)
        else {
            return AppConfig(appName: "Project Intelligence Hub", streamlitUrl: placeholderUrl)
        }
        return config
    }

    var configuredURL: URL? {
        guard streamlitUrl != Self.placeholderUrl,
              let url = URL(string: streamlitUrl),
              ["https", "http"].contains(url.scheme?.lowercased() ?? "") else {
            return nil
        }
        return url
    }
}
