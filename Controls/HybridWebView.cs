using System;
using System.Threading.Tasks;
using Microsoft.Maui.Controls;

#if WINDOWS
using Microsoft.Maui.Handlers;
using Microsoft.UI.Xaml.Controls;
using Microsoft.Web.WebView2.Core;
#endif

#if ANDROID
using Android.Runtime;
using Android.Webkit;
using Java.Interop; // <-- Add this using directive for ExportAttribute
using Microsoft.Maui.Handlers;
using AWebView = Android.Webkit.WebView;
#endif

#if IOS || MACCATALYST
using Foundation;
using WebKit;
using Microsoft.Maui.Handlers;
#endif

namespace WordFormFramework.Controls;

public partial class HybridWebView : Microsoft.Maui.Controls.WebView
{
    public event EventHandler<string>? MessageReceived;
    // Alias to match existing usage in WordFormView
    public event EventHandler<string>? ReceivedMessage;

    internal void OnMessageReceived(string message)
    {
        MessageReceived?.Invoke(this, message);
        ReceivedMessage?.Invoke(this, message);
    }

    // Helper to load raw HTML and await navigation complete (used by WordFormView)
    public Task LoadHtmlStringAsync(string html, string? baseUrl = null)
    {
        var tcs = new TaskCompletionSource<bool>();
        EventHandler<WebNavigatedEventArgs>? handler = null;
        handler = (s, e) =>
        {
            Navigated -= handler;
            tcs.TrySetResult(true);
        };
        Navigated += handler;
        Source = new HtmlWebViewSource { Html = html, BaseUrl = baseUrl };
        return tcs.Task;
    }

    static HybridWebView()
    {
#if WINDOWS
        WebViewHandler.Mapper.AppendToMapping("HybridWebViewBridge", (handler, view) =>
        {
            if (view is not HybridWebView hybrid || handler.PlatformView is not WebView2 wv2)
                return;

            void Setup(CoreWebView2 core)
            {
                core.Settings.AreDefaultContextMenusEnabled = false;
                core.WebMessageReceived -= OnWebMessage;
                core.WebMessageReceived += OnWebMessage;
            }

            void OnInitialized(object? sender, CoreWebView2InitializedEventArgs args)
            {
                if (sender is WebView2 s && s.CoreWebView2 is not null)
                    Setup(s.CoreWebView2);
            }

            void OnWebMessage(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
            {
                try
                {
                    var msg = e.TryGetWebMessageAsString();
                    if (msg is not null) hybrid.OnMessageReceived(msg);
                }
                catch { }
            }

            if (wv2.CoreWebView2 is not null)
                Setup(wv2.CoreWebView2);
            else
                wv2.CoreWebView2Initialized += OnInitialized;
        });
#endif

#if ANDROID
        WebViewHandler.Mapper.AppendToMapping("HybridWebViewBridge", (handler, view) =>
        {
            if (view is not HybridWebView hybrid || handler.PlatformView is not AWebView wv)
                return;

            wv.Settings.JavaScriptEnabled = true;
            wv.Settings.DomStorageEnabled = true;
            wv.SetOnLongClickListener(new LongClickBlocker());
            wv.LongClickable = false;

            try { wv.RemoveJavascriptInterface("native"); } catch { }
            wv.AddJavascriptInterface(new AndroidJsBridge(hybrid), "native");
        });
#endif

#if IOS || MACCATALYST
        WebViewHandler.Mapper.AppendToMapping("HybridWebViewBridge", (handler, view) =>
        {
            if (view is not HybridWebView hybrid || handler.PlatformView is not WKWebView wk)
                return;

            var controller = wk.Configuration.UserContentController;
            try { controller.RemoveScriptMessageHandler("invokeAction"); } catch { }
            controller.AddScriptMessageHandler(new IosMessageHandler(hybrid), "invokeAction");
        });
#endif
    }

#if ANDROID
    sealed class LongClickBlocker : Java.Lang.Object, Android.Views.View.IOnLongClickListener
    {
        public bool OnLongClick(Android.Views.View? v) => true;
    }

    sealed class AndroidJsBridge : Java.Lang.Object
    {
        readonly WeakReference<HybridWebView> _ref;
        public AndroidJsBridge(HybridWebView view) => _ref = new(view);

        [JavascriptInterface, Export("postMessage")]
        public void PostMessage(string message)
        {
            if (_ref.TryGetTarget(out var v))
                Microsoft.Maui.ApplicationModel.MainThread.BeginInvokeOnMainThread(() => v.OnMessageReceived(message));
        }
    }
#endif

#if IOS || MACCATALYST
    sealed class IosMessageHandler : NSObject, IWKScriptMessageHandler
    {
        readonly WeakReference<HybridWebView> _ref;
        public IosMessageHandler(HybridWebView view) => _ref = new(view);

        public void DidReceiveScriptMessage(WKUserContentController userContentController, WKScriptMessage message)
        {
            var payload = message.Body?.ToString() ?? string.Empty;
            if (_ref.TryGetTarget(out var v))
                Microsoft.Maui.ApplicationModel.MainThread.BeginInvokeOnMainThread(() => v.OnMessageReceived(payload));
        }
    }
#endif
}
