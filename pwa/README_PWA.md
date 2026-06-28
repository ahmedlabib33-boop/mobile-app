# PWA Option

This folder contains a simple manifest and service worker for a home-screen install option.

Streamlit does not automatically serve this folder as a PWA. To use it, deploy the Streamlit app behind a small static/web server or reverse proxy that also serves:

- `/manifest.json`
- `/service-worker.js`
- `/icons/icon-192.png`
- `/icons/icon-512.png`

Then add these tags to the hosted page template or proxy-injected head:

```html
<link rel="manifest" href="/manifest.json">
<meta name="theme-color" content="#0f766e">
<script>
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('/service-worker.js');
  }
</script>
```

## Director iPhone Installation Steps

This is the immediate free iPhone working route:

1. Open the deployed Streamlit URL in Safari: `https://mobile-app.streamlit.app`
2. Tap Share.
3. Tap Add to Home Screen.
4. Confirm the name Project Intelligence Hub.
5. Open the Project Intelligence Hub icon.
6. Sign in with the director username and password.

Mobile users can then open the deployed HTTPS Streamlit URL in Chrome or Safari and choose Add to Home Screen.

This is an iPhone PWA app, not an IPA. Native iOS IPA delivery requires macOS, Xcode, and Apple signing.






