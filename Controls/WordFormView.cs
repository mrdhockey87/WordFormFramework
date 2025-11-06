using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Maui.Controls;
using Microsoft.Maui.Devices;

#if WINDOWS
using Windows.Storage;
using Windows.Storage.Pickers;
using Microsoft.UI.Xaml;
using Microsoft.Maui.Platform;
using WinRT.Interop;
#endif

namespace WordFormFramework.Controls;

public partial class WordFormView : ContentView
{
    readonly HybridWebView _webView;
    public event EventHandler? ImportCompleted;
    public event EventHandler? ExportCompleted;
    public event EventHandler? RtfExportCompleted;
    public event EventHandler<Exception>? ErrorOccurred;
    private bool _isWebViewReady;

    public WordFormView()
    {
        _webView = new HybridWebView
        {
            Source = new HtmlWebViewSource { Html = BuildHtml() },
            HorizontalOptions = LayoutOptions.Fill,
            VerticalOptions = LayoutOptions.Fill
        };

        _webView.Navigated += (s, e) =>
        {
            _isWebViewReady = true;
            System.Diagnostics.Debug.WriteLine("WebView HTML loaded.");
        };

        _webView.MessageReceived += async (s, message) =>
        {
            try
            {
                if (string.IsNullOrWhiteSpace(message)) return;
                using var doc = JsonDocument.Parse(message);
                var root = doc.RootElement;
                if (root.ValueKind == JsonValueKind.Object &&
                    root.TryGetProperty("type", out var t) &&
                    string.Equals(t.GetString(), "saveImage", StringComparison.OrdinalIgnoreCase))
                {
                    var dataUrl = root.TryGetProperty("dataUrl", out var d) ? d.GetString() : null;
                    if (!string.IsNullOrWhiteSpace(dataUrl))
                        await HandleSaveImageAsync(dataUrl!);
                }
            }
            catch { }
        };

        Content = _webView;
    }

    public async Task OpenDocxFileAsync(string path)
    {
        if (!File.Exists(path)) throw new FileNotFoundException("File not found.", path);
        var bytes = await File.ReadAllBytesAsync(path);
        await WaitWebViewReadyAsync();
        await LoadDocxAsync(bytes);
    }

    public async Task OpenRtfFileAsync(string path)
    {
        if (!File.Exists(path)) throw new FileNotFoundException("File not found.", path);
        var bytes = await File.ReadAllBytesAsync(path);
        await WaitWebViewReadyAsync();
        await LoadRtfAsync(bytes);
    }

    // Legacy convenience: now no difference (protected text removed)
    public async Task OpenRtfFileAsyncNoPrompt(string path) => await OpenRtfFileAsync(path);
    public async Task OpenRtfFileAsyncPrompt(string path) => await OpenRtfFileAsync(path);

    private async Task WaitWebViewReadyAsync()
    {
        int timeout = 0;
        while (!_isWebViewReady && timeout < 50)
        {
            await Task.Delay(100);
            timeout++;
        }
        if (!_isWebViewReady) throw new InvalidOperationException("WebView is not ready yet.");
    }

    public async Task<bool> SaveDocxToFileAsync(string path)
    {
        try
        {
            var bytes = await GetDocxAsync();
            if (bytes == null) return false;
            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            await File.WriteAllBytesAsync(path, bytes);
            ExportCompleted?.Invoke(this, EventArgs.Empty);
            return true;
        }
        catch (Exception ex)
        {
            ErrorOccurred?.Invoke(this, ex);
            return false;
        }
    }

    public async Task<bool> SaveRtfToFileAsync(string path)
    {
        try
        {
            var bytes = await GetRtfAsync();
            if (bytes == null) return false;
            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            await File.WriteAllBytesAsync(path, bytes);
            RtfExportCompleted?.Invoke(this, EventArgs.Empty);
            return true;
        }
        catch (Exception ex)
        {
            ErrorOccurred?.Invoke(this, ex);
            return false;
        }
    }

    public async Task<bool> SaveDocxWithPickerAsync()
    {
        try
        {
            var bytes = await GetDocxAsync();
            if (bytes == null || bytes.Length ==0) return false;
#if WINDOWS
            var picker = new FileSavePicker
            {
                SuggestedFileName = "document",
                SuggestedStartLocation = PickerLocationId.DocumentsLibrary
            };
            picker.FileTypeChoices.Add("Word Document", new System.Collections.Generic.List<string> { ".docx" });
            var hwnd = GetWindowHandle();
            if (hwnd != IntPtr.Zero) InitializeWithWindow.Initialize(picker, hwnd);
            var file = await picker.PickSaveFileAsync();
            if (file is null) return false;
            await FileIO.WriteBytesAsync(file, bytes);
            ExportCompleted?.Invoke(this, EventArgs.Empty);
            return true;
#else
            var path = Path.Combine(FileSystem.CacheDirectory, $"document_{DateTime.Now:yyyyMMdd_HHmmss}.docx");
            await File.WriteAllBytesAsync(path, bytes);
            ExportCompleted?.Invoke(this, EventArgs.Empty);
            return true;
#endif
        }
        catch (Exception ex) { ErrorOccurred?.Invoke(this, ex); return false; }
    }

    public async Task<bool> SaveRtfWithPickerAsync()
    {
        try
        {
            var bytes = await GetRtfAsync();
            if (bytes == null || bytes.Length ==0) return false;
#if WINDOWS
            var picker = new FileSavePicker
            {
                SuggestedFileName = "document",
                SuggestedStartLocation = PickerLocationId.DocumentsLibrary
            };
            picker.FileTypeChoices.Add("Rich Text Format", new System.Collections.Generic.List<string> { ".rtf" });
            var hwnd = GetWindowHandle();
            if (hwnd != IntPtr.Zero) InitializeWithWindow.Initialize(picker, hwnd);
            var file = await picker.PickSaveFileAsync();
            if (file is null) return false;
            await FileIO.WriteBytesAsync(file, bytes);
            RtfExportCompleted?.Invoke(this, EventArgs.Empty);
            return true;
#else
            var path = Path.Combine(FileSystem.CacheDirectory, $"document_{DateTime.Now:yyyyMMdd_HHmmss}.rtf");
            await File.WriteAllBytesAsync(path, bytes);
            RtfExportCompleted?.Invoke(this, EventArgs.Empty);
            return true;
#endif
        }
        catch (Exception ex) { ErrorOccurred?.Invoke(this, ex); return false; }
    }

    public async Task LoadDocxAsync(byte[] docxBytes)
    {
        try
        {
            await WaitWebViewReadyAsync();
            string base64 = Convert.ToBase64String(docxBytes);
            string js = $"window.importDocxFromBase64('{EscapeJs(base64)}')";
            await _webView.EvaluateJavaScriptAsync(js);
            ImportCompleted?.Invoke(this, EventArgs.Empty);
        }
        catch (Exception ex) { ErrorOccurred?.Invoke(this, ex); }
    }

    public async Task LoadRtfAsync(byte[] rtfBytes)
    {
        try
        {
            await WaitWebViewReadyAsync();
            string base64 = Convert.ToBase64String(rtfBytes);
            string js = $"window.importRtfFromBase64('{EscapeJs(base64)}')";
            await _webView.EvaluateJavaScriptAsync(js);
            ImportCompleted?.Invoke(this, EventArgs.Empty);
        }
        catch (Exception ex) { ErrorOccurred?.Invoke(this, ex); }
    }

    public async Task<byte[]?> GetDocxAsync()
    {
        try
        {
            var base64 = await _webView.EvaluateJavaScriptAsync("exportDocx()");
            base64 = UnwrapJs(base64);
            if (string.IsNullOrWhiteSpace(base64)) return null;
            return Convert.FromBase64String(base64);
        }
        catch (Exception ex) { ErrorOccurred?.Invoke(this, ex); return null; }
    }

    public async Task<byte[]?> GetRtfAsync()
    {
        var base64 = await _webView.EvaluateJavaScriptAsync("exportRtf()");
        base64 = UnwrapJs(base64);
        if (string.IsNullOrWhiteSpace(base64)) return null;
        try { return Convert.FromBase64String(base64); } catch { return null; }
    }

    static string EscapeJs(string s) => s.Replace("\\", "\\\\").Replace("'", "\\'").Replace("\r", "\\r").Replace("\n", "\\n");
    static string UnwrapJs(string s)
    {
        if (string.IsNullOrEmpty(s)) return s;
        if (s.Length >= 2 && s.StartsWith("\"") && s.EndsWith("\""))
            s = s.Substring(1, s.Length - 2).Replace("\\n", "\n").Replace("\\r", "\r").Replace("\\\"", "\"");
        return s;
    }

    static string BuildHtml()
    {
        string GetRes(string name)
        {
            var assembly = typeof(WordFormView).Assembly;
            var resourceName = assembly.GetManifestResourceNames()
                .FirstOrDefault(r => r.EndsWith(name, StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidOperationException($"Embedded resource '{name}' not found.");
            using var stream = assembly.GetManifestResourceStream(resourceName)!;
            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }
        string GetResOptional(string name)
        {
            try
            {
                var assembly = typeof(WordFormView).Assembly;
                var resourceName = assembly.GetManifestResourceNames()
                    .FirstOrDefault(r => r.EndsWith(name, StringComparison.OrdinalIgnoreCase));
                if (resourceName is null) return string.Empty;
                using var stream = assembly.GetManifestResourceStream(resourceName)!;
                using var reader = new StreamReader(stream);
                return reader.ReadToEnd();
            }
            catch { return string.Empty; }
        }

        var quillCss = GetRes("quill.snow.css");
        var quillJs = GetRes("quill.min.js");
        var mammothJs = GetRes("mammoth.browser.min.js");
        var htmlDocxJs = GetRes("html-docx.js");
        var rtfToHtmlJs = GetResOptional("rtf-to-html.min.js");

        var disableContextMenu = DeviceInfo.Current.Platform == DevicePlatform.MacCatalyst;
        var disableFlag = disableContextMenu ? "true" : "false";

        var sb = new StringBuilder();
        sb.AppendLine("<!doctype html>");
        sb.AppendLine("<html>");
        sb.AppendLine("<head>");
        sb.AppendLine("<meta charset='utf-8'>");
        sb.Append("<style>").Append(quillCss).AppendLine("</style>");
        sb.AppendLine(@"<style>");
        sb.AppendLine(@"html,body{height:100%;margin:0;padding:0;background:#ffffff;}");
        sb.AppendLine(@"body{display:flex;flex-direction:column;}");
        sb.AppendLine(@"#editor{flex:1;min-height:0;}");
        sb.AppendLine(@".wf-img-menu{position:fixed;z-index:9999;background:#fff;border:1px solid #ccc;box-shadow:02px6px rgba(0,0,0,.25);border-radius:4px;font:14px -apple-system,Segoe UI,Arial,sans-serif;min-width:150px;padding:4px;display:none;}");
        sb.AppendLine(@".wf-img-menu button{all:unset;display:block;width:100%;padding:6px10px;cursor:pointer;border-radius:3px;}");
        sb.AppendLine(@".wf-img-menu button:hover{background:#e6f0ff;}");
        sb.AppendLine("</style>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");

        sb.AppendLine("<div id='editor'></div>");
        sb.AppendLine("<div class='wf-img-menu' id='wfImgMenu'><button data-action='save'>Save Image...</button></div>");

        // assets
        sb.Append("<script>").Append(quillJs).AppendLine("</script>");
        sb.Append("<script>").Append(mammothJs).AppendLine("</script>");
        sb.Append("<script>").Append(htmlDocxJs).AppendLine("</script>");
        if (!string.IsNullOrWhiteSpace(rtfToHtmlJs))
            sb.Append("<script>").Append(rtfToHtmlJs).AppendLine("</script>");

        // app script
        sb.AppendLine("<script>");
        sb.AppendLine(@"function nativePostMessage(message){try{if(window.chrome?.webview?.postMessage)window.chrome.webview.postMessage(message);else if(window.webkit?.messageHandlers?.invokeAction)window.webkit.messageHandlers.invokeAction.postMessage(message);else if(window.native?.postMessage)window.native.postMessage(message);}catch{}}");
        sb.Append("const DISABLE_CONTEXT_MENU = ").Append(disableFlag).AppendLine(";");

        sb.AppendLine(@"if(DISABLE_CONTEXT_MENU){window.addEventListener('contextmenu',e=>e.preventDefault(),false);} ");

        // Quill init with full toolbar (auto-generated)
        sb.AppendLine(@"var quill;(function(){");
        sb.AppendLine(@" var toolbarOptions = [[{ header: [1,2,3,false] }], ['bold','italic','underline','strike'], [{ 'color': [] }, { 'background': [] }], [{ 'list': 'ordered' }, { 'list': 'bullet' }], [{ 'align': [] }], ['link','image'], ['blockquote','code-block'], ['clean']];");
        sb.AppendLine(@" quill = new Quill('#editor',{ theme:'snow', modules:{ toolbar: toolbarOptions } });");
        sb.AppendLine(@"})();");

        // Image context menu
        sb.AppendLine(@"const menu=document.getElementById('wfImgMenu'); function hideMenu(){menu.style.display='none';document.removeEventListener('click',onDocClick,true);} function onDocClick(){hideMenu();} function showMenu(x,y,img){menu.style.display='block'; const W=innerWidth,H=innerHeight,r=menu.getBoundingClientRect(); if(x+r.width>W)x=W-r.width-4; if(y+r.height>H)y=H-r.height-4; menu.style.left=x+'px'; menu.style.top=y+'px'; document.addEventListener('click',onDocClick,true); menu.onclick=ev=>{const btn=ev.target.closest('button'); if(!btn)return; if(btn.getAttribute('data-action')==='save'){ try{ const c=document.createElement('canvas'); c.width=img.naturalWidth||img.width; c.height=img.naturalHeight||img.height; c.getContext('2d').drawImage(img,0,0); const dataUrl=c.toDataURL('image/png'); nativePostMessage(JSON.stringify({type:'saveImage',dataUrl})); }catch{} hideMenu(); }}; }");
        sb.AppendLine(@"document.addEventListener('contextmenu',e=>{const t=e.target; if(t && t.tagName==='IMG'){ e.preventDefault(); showMenu(e.clientX,e.clientY,t);} },false);");

        // DOCX import
        sb.AppendLine(@"window.importDocxFromBase64=async function(b64){ try{ var bin=atob(b64); var bytes=new Uint8Array(bin.length); for(var i=0;i<bin.length;i++) bytes[i]=bin.charCodeAt(i); var result=await mammoth.convertToHtml({arrayBuffer:bytes.buffer}); quill.root.innerHTML=result.value||'<p><em>No content</em></p>'; return true;}catch(e){return false;} };");

        // RTF import fallback
        sb.AppendLine(@"window.importRtfFromBase64=async function(b64){ try{ var rtf=atob(b64); var plain=rtf.replace(/\'[0-9a-fA-F]{2}/g,' ').replace(/\par/gi,'\n').replace(/\[a-zA-Z]+-?\d* ?/g,'').replace(/[{}]/g,'').trim(); if(!plain) plain='(empty)'; var html='<p>'+plain.replace(/\n+/g,'</p><p>')+'</p>'; quill.root.innerHTML=html; return true;}catch(e){return false;} };");

        // export helpers
        sb.AppendLine(@"function blobToBase64(blob){return new Promise((res,rej)=>{var fr=new FileReader(); fr.onloadend=()=>res(fr.result.split(',')[1]); fr.onerror=rej; fr.readAsDataURL(blob);});}");
        sb.AppendLine(@"window.exportDocx=async function(){ try{ var html='<html><body>'+quill.root.innerHTML+'</body></html>'; var blob=htmlDocx.asBlob(html); return await blobToBase64(blob);}catch(e){return '';} };");
        sb.AppendLine(@"window.exportRtf=function(){ try{ var txt=quill.getText(); txt=txt.replace(/\\/g,'\\\\').replace(/\{/g,'\\{').replace(/\}/g,'\\}').replace(/\r?\n/g,'\\par '); var rtf='{\\rtf1\\ansi\\deff0 '+txt+'}'; return btoa(unescape(encodeURIComponent(rtf))); }catch(e){ return ''; } };");

        sb.AppendLine("</script>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        return sb.ToString();
    }

    private static (string Mime, byte[] Bytes) ParseDataUrl(string dataUrl)
    {
        var comma = dataUrl.IndexOf(',');
        if (comma < 0) return ("application/octet-stream", Array.Empty<byte>());
        var meta = dataUrl.Substring(5, comma - 5);
        var base64 = dataUrl[(comma + 1)..];
        var semi = meta.IndexOf(';');
        var mime = semi >= 0 ? meta.Substring(0, semi) : meta;
        return (mime, Convert.FromBase64String(base64));
    }

    private async Task HandleSaveImageAsync(string dataUrl)
    {
        try
        {
            var (mime, bytes) = ParseDataUrl(dataUrl);
            if (bytes.Length == 0) return;
#if WINDOWS
            var picker = new FileSavePicker { SuggestedFileName = "image", SuggestedStartLocation = PickerLocationId.PicturesLibrary };
            picker.FileTypeChoices.Add("PNG Image", new System.Collections.Generic.List<string> { ".png" });
            var hwnd = GetWindowHandle();
            if (hwnd != IntPtr.Zero) InitializeWithWindow.Initialize(picker, hwnd);
            var file = await picker.PickSaveFileAsync();
            if (file is not null) await FileIO.WriteBytesAsync(file, bytes);
#else
            var ext = mime.EndsWith("png", StringComparison.OrdinalIgnoreCase) ? ".png" : ".bin";
            var path = Path.Combine(FileSystem.CacheDirectory, $"image_{DateTime.Now:yyyyMMdd_HHmmss}{ext}");
            await File.WriteAllBytesAsync(path, bytes);
#endif
        }
        catch (Exception ex) { ErrorOccurred?.Invoke(this, ex); }
    }

#if WINDOWS
    private static IntPtr GetWindowHandle()
    {
        try
        {
            var mauiWindow = Microsoft.Maui.Controls.Application.Current?.Windows?.FirstOrDefault();
            var nativeWindow = mauiWindow?.Handler?.PlatformView as Microsoft.UI.Xaml.Window;
            if (nativeWindow is not null) return WindowNative.GetWindowHandle(nativeWindow);
        }
        catch { }
        return IntPtr.Zero;
    }
#endif
}