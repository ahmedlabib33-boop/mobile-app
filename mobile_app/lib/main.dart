import 'dart:convert';

import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:url_launcher/url_launcher.dart';
import 'package:webview_flutter/webview_flutter.dart';

void main() {
  WidgetsFlutterBinding.ensureInitialized();
  runApp(const ProjectIntelligenceHubApp());
}

class MobileConfig {
  const MobileConfig({
    required this.appName,
    required this.streamlitUrl,
  });

  final String appName;
  final String streamlitUrl;

  bool get hasConfiguredUrl =>
      streamlitUrl.trim().isNotEmpty &&
      streamlitUrl.trim() != 'PUT_DEPLOYED_STREAMLIT_URL_HERE' &&
      Uri.tryParse(streamlitUrl.trim())?.hasScheme == true &&
      (Uri.tryParse(streamlitUrl.trim())?.host.isNotEmpty ?? false);

  static Future<MobileConfig> load() async {
    final raw = await rootBundle.loadString('assets/mobile_config.json');
    final decoded = jsonDecode(raw) as Map<String, dynamic>;
    return MobileConfig(
      appName: (decoded['app_name'] as String?) ?? 'Project Intelligence Hub',
      streamlitUrl:
          (decoded['streamlit_url'] as String?) ?? 'PUT_DEPLOYED_STREAMLIT_URL_HERE',
    );
  }
}

class ProjectIntelligenceHubApp extends StatelessWidget {
  const ProjectIntelligenceHubApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Project Intelligence Hub',
      debugShowCheckedModeBanner: false,
      themeMode: ThemeMode.system,
      theme: _buildTheme(Brightness.light),
      darkTheme: _buildTheme(Brightness.dark),
      home: const MobileShell(),
    );
  }

  ThemeData _buildTheme(Brightness brightness) {
    final isDark = brightness == Brightness.dark;
    final scheme = ColorScheme.fromSeed(
      seedColor: const Color(0xFF0F766E),
      brightness: brightness,
    );
    return ThemeData(
      useMaterial3: true,
      colorScheme: scheme,
      scaffoldBackgroundColor:
          isDark ? const Color(0xFF07111F) : const Color(0xFFF5F7FB),
      appBarTheme: AppBarTheme(
        centerTitle: false,
        elevation: 0,
        backgroundColor: isDark ? const Color(0xFF0B1220) : Colors.white,
        foregroundColor: isDark ? Colors.white : const Color(0xFF111827),
        surfaceTintColor: Colors.transparent,
      ),
    );
  }
}

class MobileShell extends StatefulWidget {
  const MobileShell({super.key});

  @override
  State<MobileShell> createState() => _MobileShellState();
}

class _MobileShellState extends State<MobileShell> {
  WebViewController? _controller;
  MobileConfig? _config;
  bool _loading = true;
  bool _hasPageError = false;
  String _statusText = 'Preparing secure mobile workspace';
  double _progress = 0;

  @override
  void initState() {
    super.initState();
    _prepare();
  }

  Future<void> _prepare() async {
    final config = await MobileConfig.load();
    if (!config.hasConfiguredUrl) {
      setState(() {
        _config = config;
        _loading = false;
        _hasPageError = true;
        _statusText = 'Streamlit URL is not configured yet.';
      });
      return;
    }

    final controller = WebViewController()
      ..setJavaScriptMode(JavaScriptMode.unrestricted)
      ..setBackgroundColor(Colors.transparent)
      ..setNavigationDelegate(
        NavigationDelegate(
          onProgress: (progress) {
            setState(() {
              _progress = progress / 100;
              _loading = progress < 100;
              _statusText = progress < 100
                  ? 'Loading Project Intelligence Hub'
                  : 'Project Intelligence Hub is ready';
            });
          },
          onPageStarted: (_) {
            setState(() {
              _loading = true;
              _hasPageError = false;
              _statusText = 'Loading Project Intelligence Hub';
            });
          },
          onPageFinished: (_) {
            setState(() {
              _loading = false;
              _statusText = 'Project Intelligence Hub is ready';
            });
          },
          onWebResourceError: (error) {
            if (error.isForMainFrame == true) {
              setState(() {
                _loading = false;
                _hasPageError = true;
                _statusText = 'Unable to reach the deployed dashboard.';
              });
            }
          },
          onNavigationRequest: _handleNavigation,
        ),
      );

    await controller.loadRequest(Uri.parse(config.streamlitUrl.trim()));
    setState(() {
      _config = config;
      _controller = controller;
    });
  }

  NavigationDecision _handleNavigation(NavigationRequest request) {
    final uri = Uri.tryParse(request.url);
    if (uri == null) {
      return NavigationDecision.prevent;
    }

    final isConfiguredHost =
        Uri.parse(_config!.streamlitUrl.trim()).host.toLowerCase() ==
            uri.host.toLowerCase();
    final isExport = _looksLikeDownload(uri);

    if (!isConfiguredHost || isExport) {
      _openExternal(uri);
      return NavigationDecision.prevent;
    }

    return NavigationDecision.navigate;
  }

  bool _looksLikeDownload(Uri uri) {
    final path = uri.path.toLowerCase();
    return path.endsWith('.pdf') ||
        path.endsWith('.xlsx') ||
        path.endsWith('.xls') ||
        path.endsWith('.csv') ||
        path.endsWith('.docx') ||
        path.endsWith('.pptx') ||
        path.endsWith('.zip') ||
        path.contains('download');
  }

  Future<void> _openExternal(Uri uri) async {
    await launchUrl(uri, mode: LaunchMode.externalApplication);
  }

  Future<void> _refresh() async {
    setState(() {
      _hasPageError = false;
      _loading = true;
    });
    await _controller?.reload();
  }

  Future<bool> _handleBack() async {
    final controller = _controller;
    if (controller != null && await controller.canGoBack()) {
      await controller.goBack();
      return false;
    }
    return true;
  }

  @override
  Widget build(BuildContext context) {
    return PopScope(
      canPop: false,
      onPopInvoked: (didPop) async {
        if (didPop) {
          return;
        }
        final shouldClose = await _handleBack();
        if (shouldClose && context.mounted) {
          SystemNavigator.pop();
        }
      },
      child: Scaffold(
        appBar: AppBar(
          titleSpacing: 16,
          title: const Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Text(
                'Project Intelligence Hub',
                maxLines: 1,
                overflow: TextOverflow.ellipsis,
                style: TextStyle(fontWeight: FontWeight.w800, fontSize: 17),
              ),
              Text(
                'Project controls command center',
                maxLines: 1,
                overflow: TextOverflow.ellipsis,
                style: TextStyle(fontWeight: FontWeight.w500, fontSize: 11),
              ),
            ],
          ),
          actions: [
            IconButton(
              tooltip: 'Refresh',
              onPressed: _controller == null ? null : _refresh,
              icon: const Icon(Icons.refresh_rounded),
            ),
          ],
          bottom: PreferredSize(
            preferredSize: const Size.fromHeight(3),
            child: _loading
                ? LinearProgressIndicator(value: _progress == 0 ? null : _progress)
                : const SizedBox(height: 3),
          ),
        ),
        body: SafeArea(
          child: _buildBody(),
        ),
      ),
    );
  }

  Widget _buildBody() {
    if (_config == null || _controller == null || _hasPageError) {
      return _StatusPanel(
        title: _statusText,
        message: _config?.hasConfiguredUrl == false
            ? 'Edit mobile_config.json and mobile_app/assets/mobile_config.json with the deployed Streamlit HTTPS URL, then rebuild the app.'
            : 'Check internet connectivity and confirm the deployed Streamlit URL is reachable.',
        buttonLabel: _config?.hasConfiguredUrl == false ? null : 'Try again',
        onPressed: _config?.hasConfiguredUrl == false ? null : _refresh,
      );
    }

    return Stack(
      children: [
        RefreshIndicator(
          onRefresh: _refresh,
          child: CustomScrollView(
            physics: const AlwaysScrollableScrollPhysics(),
            slivers: [
              SliverFillRemaining(
                child: WebViewWidget(controller: _controller!),
              ),
            ],
          ),
        ),
        if (_loading)
          _LoadingOverlay(statusText: _statusText),
      ],
    );
  }
}

class _LoadingOverlay extends StatelessWidget {
  const _LoadingOverlay({required this.statusText});

  final String statusText;

  @override
  Widget build(BuildContext context) {
    return DecoratedBox(
      decoration: BoxDecoration(
        color: Theme.of(context).scaffoldBackgroundColor.withOpacity(.88),
      ),
      child: Center(
        child: ConstrainedBox(
          constraints: const BoxConstraints(maxWidth: 340),
          child: Card(
            elevation: 0,
            shape: RoundedRectangleBorder(
              borderRadius: BorderRadius.circular(12),
              side: BorderSide(color: Theme.of(context).dividerColor),
            ),
            child: Padding(
              padding: const EdgeInsets.all(22),
              child: Column(
                mainAxisSize: MainAxisSize.min,
                children: [
                  const CircularProgressIndicator(),
                  const SizedBox(height: 18),
                  Text(
                    statusText,
                    textAlign: TextAlign.center,
                    style: const TextStyle(fontWeight: FontWeight.w700),
                  ),
                ],
              ),
            ),
          ),
        ),
      ),
    );
  }
}

class _StatusPanel extends StatelessWidget {
  const _StatusPanel({
    required this.title,
    required this.message,
    this.buttonLabel,
    this.onPressed,
  });

  final String title;
  final String message;
  final String? buttonLabel;
  final VoidCallback? onPressed;

  @override
  Widget build(BuildContext context) {
    return Center(
      child: Padding(
        padding: const EdgeInsets.all(24),
        child: ConstrainedBox(
          constraints: const BoxConstraints(maxWidth: 420),
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              Icon(
                Icons.signal_wifi_connected_no_internet_4_rounded,
                size: 56,
                color: Theme.of(context).colorScheme.primary,
              ),
              const SizedBox(height: 18),
              Text(
                title,
                textAlign: TextAlign.center,
                style: Theme.of(context).textTheme.titleLarge?.copyWith(
                      fontWeight: FontWeight.w800,
                    ),
              ),
              const SizedBox(height: 10),
              Text(
                message,
                textAlign: TextAlign.center,
                style: Theme.of(context).textTheme.bodyMedium,
              ),
              if (buttonLabel != null) ...[
                const SizedBox(height: 18),
                FilledButton.icon(
                  onPressed: onPressed,
                  icon: const Icon(Icons.refresh_rounded),
                  label: Text(buttonLabel!),
                ),
              ],
            ],
          ),
        ),
      ),
    );
  }
}
