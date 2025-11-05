using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Text.Json;

#if WINDOWS
using Microsoft.Maui.Handlers;
using Microsoft.UI.Xaml.Controls;
using Microsoft.Web.WebView2.Core;
using Windows.Storage.Pickers;
using WinRT.Interop;
#endif

#if MACCATALYST
using Microsoft.Maui.Handlers;
using WebKit;
#endif

namespace WordFormFramework.Demo;

public static class MauiProgram
{
	public static MauiApp CreateMauiApp()
	{
		var builder = MauiApp.CreateBuilder();
		builder
			.UseMauiApp<App>()
			.ConfigureFonts(fonts =>
			{
				fonts.AddFont("OpenSans-Regular.ttf", "OpenSansRegular");
				fonts.AddFont("OpenSans-Semibold.ttf", "OpenSansSemibold");
			})
#if WINDOWS || MACCATALYST
 .ConfigureMauiHandlers(handlers =>
 {
#if WINDOWS
 // Windows (WebView2):
 // - Keep default context menu for general use
 // - Intercept downloads via DownloadStarting (links/files)
 // - Handle image save via JS postMessage to avoid double Save dialogs
 WebViewHandler.Mapper.AppendToMapping("WinWebView2Downloads", (handler, view) =>
 {
 var native = (WebView2)handler.PlatformView;

 void Attach()
 {
 var core = native.CoreWebView2;
 if (core == null) return;

 core.Settings.AreDefaultContextMenusEnabled = true;
 core.Settings.IsWebMessageEnabled = true;

 core.DownloadStarting -= OnDownloadStarting;
 core.DownloadStarting += OnDownloadStarting;

 core.WebMessageReceived -= OnWebMessage;
 core.WebMessageReceived += OnWebMessage;
 }

 void OnCoreInitialized(object? s, CoreWebView2InitializedEventArgs e)
 {
 native.CoreWebView2Initialized -= OnCoreInitialized;
 Attach();
 }

 if (native.CoreWebView2 != null) Attach();
 else native.CoreWebView2Initialized += OnCoreInitialized;

 async void OnDownloadStarting(CoreWebView2 sender, CoreWebView2DownloadStartingEventArgs e)
 {
 // Downloads initiated by navigation or anchor links
 var deferral = e.GetDeferral();
 try
 {
 var picker = new FileSavePicker();
 var win = (Microsoft.UI.Xaml.Window)Application.Current!.Windows[0].Handler.PlatformView!;
 InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(win));

 var filename = System.IO.Path.GetFileName(e.ResultFilePath);
 if (string.IsNullOrWhiteSpace(filename)) filename = "download";
 var ext = System.IO.Path.GetExtension(filename);
 if (string.IsNullOrWhiteSpace(ext)) ext = ".bin";

 picker.SuggestedFileName = System.IO.Path.GetFileNameWithoutExtension(filename);
 picker.FileTypeChoices.Add("File", new List<string> { ext.ToLowerInvariant() });

 var file = await picker.PickSaveFileAsync();
 if (file == null)
 {
 e.Cancel = true;
 return;
 }

 e.ResultFilePath = file.Path;
 e.Handled = true; // only our dialog
 }
 catch
 {
 e.Cancel = true;
 }
 finally
 {
 deferral.Complete();
 }
 }

 async void OnWebMessage(CoreWebView2 sender, CoreWebView2WebMessageReceivedEventArgs e)
 {
 try
 {
 var text = e.TryGetWebMessageAsString();
 if (string.IsNullOrWhiteSpace(text)) return;
 using var doc = JsonDocument.Parse(text);
 var root = doc.RootElement;
 if (!root.TryGetProperty("type", out var typeEl)) return;
 if (!string.Equals(typeEl.GetString(), "saveImage", System.StringComparison.Ordinal)) return;
 if (!root.TryGetProperty("dataUrl", out var dataUrlEl)) return;
 var dataUrl = dataUrlEl.GetString();
 if (string.IsNullOrWhiteSpace(dataUrl)) return;

 var comma = dataUrl!.IndexOf(',');
 if (comma <0) return;
 var base64 = dataUrl.Substring(comma +1);

 byte[] content;
 try { content = System.Convert.FromBase64String(base64); }
 catch { return; }

 var picker = new FileSavePicker();
 var win = (Microsoft.UI.Xaml.Window)Application.Current!.Windows[0].Handler.PlatformView!;
 InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(win));
 picker.SuggestedFileName = "image";
 picker.FileTypeChoices.Add("Image", new List<string> { ".png" });

 var file = await picker.PickSaveFileAsync();
 if (file == null) return;
 await System.IO.File.WriteAllBytesAsync(file.Path, content);
 }
 catch { }
 }
 });
#endif

#if MACCATALYST
 // MacCatalyst (WKWebView): reduce native previews
 WebViewHandler.Mapper.AppendToMapping("MacConfig", (handler, view) =>
 {
 var wk = (WKWebView)handler.PlatformView;
 wk.AllowsLinkPreview = false;
 });
#endif
 })
#endif
 ;

#if DEBUG
		builder.Logging.AddDebug();
#endif

		return builder.Build();
	}
}
