package com.samco.projectintelligencehub;

import android.annotation.SuppressLint;
import android.app.Activity;
import android.app.DownloadManager;
import android.content.ActivityNotFoundException;
import android.content.Context;
import android.content.Intent;
import android.graphics.Color;
import android.net.ConnectivityManager;
import android.net.NetworkCapabilities;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.view.Gravity;
import android.view.MotionEvent;
import android.view.View;
import android.view.ViewGroup;
import android.webkit.CookieManager;
import android.webkit.DownloadListener;
import android.webkit.URLUtil;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceError;
import android.webkit.WebResourceRequest;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.Button;
import android.widget.FrameLayout;
import android.widget.LinearLayout;
import android.widget.ProgressBar;
import android.widget.TextView;
import android.widget.Toast;

import org.json.JSONObject;

import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;

public class MainActivity extends Activity {
    private static final String PLACEHOLDER_URL = "PUT_DEPLOYED_STREAMLIT_URL_HERE";

    private WebView webView;
    private ProgressBar progressBar;
    private LinearLayout statusPanel;
    private TextView statusTitle;
    private TextView statusMessage;
    private String streamlitUrl = PLACEHOLDER_URL;
    private float pullStartY = 0f;
    private boolean pullTracking = false;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        configureSystemBars();
        streamlitUrl = loadStreamlitUrlFromAssets();
        buildLayout();
        configureWebView();
        loadDashboard();
    }

    private void configureSystemBars() {
        getWindow().setStatusBarColor(Color.rgb(7, 17, 31));
        getWindow().setNavigationBarColor(Color.rgb(7, 17, 31));
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M) {
            getWindow().getDecorView().setSystemUiVisibility(0);
        }
    }

    private String loadStreamlitUrlFromAssets() {
        try (InputStream input = getAssets().open("mobile_config.json");
             BufferedReader reader = new BufferedReader(new InputStreamReader(input, StandardCharsets.UTF_8))) {
            StringBuilder builder = new StringBuilder();
            String line;
            while ((line = reader.readLine()) != null) {
                builder.append(line);
            }
            JSONObject config = new JSONObject(builder.toString());
            return config.optString("streamlit_url", PLACEHOLDER_URL).trim();
        } catch (Exception ignored) {
            return PLACEHOLDER_URL;
        }
    }

    private void buildLayout() {
        FrameLayout root = new FrameLayout(this);
        root.setBackgroundColor(Color.rgb(245, 247, 251));

        webView = new WebView(this);
        webView.setLayoutParams(new FrameLayout.LayoutParams(
                ViewGroup.LayoutParams.MATCH_PARENT,
                ViewGroup.LayoutParams.MATCH_PARENT
        ));

        progressBar = new ProgressBar(this, null, android.R.attr.progressBarStyleHorizontal);
        FrameLayout.LayoutParams progressParams = new FrameLayout.LayoutParams(
                ViewGroup.LayoutParams.MATCH_PARENT,
                dp(4)
        );
        progressParams.gravity = Gravity.TOP;
        progressBar.setMax(100);
        progressBar.setVisibility(View.GONE);

        statusPanel = new LinearLayout(this);
        statusPanel.setOrientation(LinearLayout.VERTICAL);
        statusPanel.setGravity(Gravity.CENTER);
        statusPanel.setPadding(dp(24), dp(24), dp(24), dp(24));
        statusPanel.setBackgroundColor(Color.rgb(245, 247, 251));
        statusPanel.setVisibility(View.GONE);

        statusTitle = new TextView(this);
        statusTitle.setTextColor(Color.rgb(15, 23, 42));
        statusTitle.setTextSize(22);
        statusTitle.setGravity(Gravity.CENTER);
        statusTitle.setTypeface(android.graphics.Typeface.DEFAULT_BOLD);

        statusMessage = new TextView(this);
        statusMessage.setTextColor(Color.rgb(71, 85, 105));
        statusMessage.setTextSize(15);
        statusMessage.setGravity(Gravity.CENTER);
        LinearLayout.LayoutParams messageParams = new LinearLayout.LayoutParams(
                ViewGroup.LayoutParams.MATCH_PARENT,
                ViewGroup.LayoutParams.WRAP_CONTENT
        );
        messageParams.setMargins(0, dp(12), 0, dp(20));

        Button retryButton = new Button(this);
        retryButton.setText("Try Again");
        retryButton.setAllCaps(false);
        retryButton.setOnClickListener(v -> loadDashboard());

        statusPanel.addView(statusTitle);
        statusPanel.addView(statusMessage, messageParams);
        statusPanel.addView(retryButton);

        root.addView(webView);
        root.addView(progressBar, progressParams);
        root.addView(statusPanel, new FrameLayout.LayoutParams(
                ViewGroup.LayoutParams.MATCH_PARENT,
                ViewGroup.LayoutParams.MATCH_PARENT
        ));
        setContentView(root);
    }

    @SuppressLint("SetJavaScriptEnabled")
    private void configureWebView() {
        WebSettings settings = webView.getSettings();
        settings.setJavaScriptEnabled(true);
        settings.setDomStorageEnabled(true);
        settings.setDatabaseEnabled(true);
        settings.setLoadWithOverviewMode(true);
        settings.setUseWideViewPort(true);
        settings.setBuiltInZoomControls(false);
        settings.setDisplayZoomControls(false);
        settings.setAllowFileAccess(false);
        settings.setAllowContentAccess(true);
        settings.setMediaPlaybackRequiresUserGesture(false);
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.O) {
            settings.setSafeBrowsingEnabled(true);
        }

        CookieManager.getInstance().setAcceptCookie(true);
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.LOLLIPOP) {
            CookieManager.getInstance().setAcceptThirdPartyCookies(webView, true);
        }

        webView.setWebChromeClient(new WebChromeClient() {
            @Override
            public void onProgressChanged(WebView view, int newProgress) {
                progressBar.setVisibility(newProgress >= 100 ? View.GONE : View.VISIBLE);
                progressBar.setProgress(newProgress);
            }
        });

        webView.setWebViewClient(new WebViewClient() {
            @Override
            public boolean shouldOverrideUrlLoading(WebView view, WebResourceRequest request) {
                return handleNavigation(request.getUrl());
            }

            @Override
            public boolean shouldOverrideUrlLoading(WebView view, String url) {
                return handleNavigation(Uri.parse(url));
            }

            @Override
            public void onPageFinished(WebView view, String url) {
                statusPanel.setVisibility(View.GONE);
                webView.setVisibility(View.VISIBLE);
            }

            @Override
            public void onReceivedError(WebView view, WebResourceRequest request, WebResourceError error) {
                if (Build.VERSION.SDK_INT < Build.VERSION_CODES.M || request.isForMainFrame()) {
                    showStatus(
                            "Dashboard unavailable",
                            "Check the internet connection and confirm the deployed Streamlit URL is reachable."
                    );
                }
            }
        });

        webView.setDownloadListener(buildDownloadListener());
        webView.setOnTouchListener((view, event) -> {
            if (event.getAction() == MotionEvent.ACTION_DOWN && webView.getScrollY() == 0) {
                pullStartY = event.getY();
                pullTracking = true;
            } else if (event.getAction() == MotionEvent.ACTION_UP && pullTracking) {
                float distance = event.getY() - pullStartY;
                pullTracking = false;
                if (webView.getScrollY() == 0 && distance > dp(96)) {
                    Toast.makeText(this, "Refreshing dashboard", Toast.LENGTH_SHORT).show();
                    webView.reload();
                }
            } else if (event.getAction() == MotionEvent.ACTION_CANCEL) {
                pullTracking = false;
            }
            return false;
        });
    }

    private DownloadListener buildDownloadListener() {
        return (url, userAgent, contentDisposition, mimeType, contentLength) -> {
            try {
                DownloadManager.Request request = new DownloadManager.Request(Uri.parse(url));
                String fileName = URLUtil.guessFileName(url, contentDisposition, mimeType);
                request.setMimeType(mimeType);
                request.addRequestHeader("User-Agent", userAgent);
                request.setTitle(fileName);
                request.setDescription("Downloading Project Intelligence Hub export");
                request.setNotificationVisibility(DownloadManager.Request.VISIBILITY_VISIBLE_NOTIFY_COMPLETED);
                request.setDestinationInExternalPublicDir(Environment.DIRECTORY_DOWNLOADS, fileName);
                DownloadManager manager = (DownloadManager) getSystemService(DOWNLOAD_SERVICE);
                manager.enqueue(request);
                Toast.makeText(this, "Download started: " + fileName, Toast.LENGTH_LONG).show();
            } catch (Exception ex) {
                openExternal(Uri.parse(url));
            }
        };
    }

    private boolean handleNavigation(Uri uri) {
        if (uri == null) {
            return true;
        }
        Uri configured = Uri.parse(streamlitUrl);
        boolean sameHost = configured.getHost() != null
                && configured.getHost().equalsIgnoreCase(uri.getHost());
        boolean exportLink = looksLikeExport(uri);
        if (!sameHost || exportLink) {
            openExternal(uri);
            return true;
        }
        return false;
    }

    private boolean looksLikeExport(Uri uri) {
        String path = uri.getPath() == null ? "" : uri.getPath().toLowerCase();
        return path.endsWith(".pdf")
                || path.endsWith(".xlsx")
                || path.endsWith(".xls")
                || path.endsWith(".csv")
                || path.endsWith(".docx")
                || path.endsWith(".pptx")
                || path.endsWith(".zip")
                || path.contains("download");
    }

    private void openExternal(Uri uri) {
        try {
            startActivity(new Intent(Intent.ACTION_VIEW, uri));
        } catch (ActivityNotFoundException ex) {
            Toast.makeText(this, "No app can open this link.", Toast.LENGTH_LONG).show();
        }
    }

    private void loadDashboard() {
        if (!isConfiguredProductionUrl()) {
            showStatus(
                "Streamlit URL not configured",
                    "Edit android_app/app/src/main/assets/mobile_config.json and replace the placeholder with the deployed HTTPS Streamlit URL."
            );
            return;
        }
        if (!hasInternetConnection()) {
            showStatus(
                    "No internet connection",
                    "Connect to the internet and try again."
            );
            return;
        }
        statusPanel.setVisibility(View.GONE);
        webView.setVisibility(View.VISIBLE);
        webView.loadUrl(streamlitUrl);
    }

    private boolean isConfiguredProductionUrl() {
        return streamlitUrl != null
                && !streamlitUrl.isEmpty()
                && !PLACEHOLDER_URL.equals(streamlitUrl)
                && streamlitUrl.startsWith("https://");
    }

    private boolean hasInternetConnection() {
        ConnectivityManager manager = (ConnectivityManager) getSystemService(Context.CONNECTIVITY_SERVICE);
        if (manager == null) {
            return true;
        }
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M) {
            NetworkCapabilities capabilities = manager.getNetworkCapabilities(manager.getActiveNetwork());
            return capabilities != null
                    && capabilities.hasCapability(NetworkCapabilities.NET_CAPABILITY_INTERNET);
        }
        return manager.getActiveNetworkInfo() != null && manager.getActiveNetworkInfo().isConnected();
    }

    private void showStatus(String title, String message) {
        webView.setVisibility(View.GONE);
        progressBar.setVisibility(View.GONE);
        statusTitle.setText(title);
        statusMessage.setText(message);
        statusPanel.setVisibility(View.VISIBLE);
    }

    @Override
    public void onBackPressed() {
        if (webView != null && webView.canGoBack()) {
            webView.goBack();
            return;
        }
        super.onBackPressed();
    }

    private int dp(int value) {
        return Math.round(value * getResources().getDisplayMetrics().density);
    }
}
