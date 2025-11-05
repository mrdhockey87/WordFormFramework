using System;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using Microsoft.Maui.Controls;
using Microsoft.Maui.Devices;

namespace WordFormFramework.Controls;

public partial class WordFormView : ContentView
{
    readonly WebView _webView;
    public event EventHandler? ImportCompleted;
    public event EventHandler? ExportCompleted;
    public event EventHandler<Exception>? ErrorOccurred;
    private bool _isWebViewReady;

    public WordFormView()
    {
        _webView = new WebView
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

        Content = _webView;
    }

    public async Task OpenDocxFileAsync(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("File not found.", path);

        var bytes = await File.ReadAllBytesAsync(path);

        int timeout =0;
        while (!_isWebViewReady && timeout <50)
        {
            await Task.Delay(100);
            timeout++;
        }

        await LoadDocxAsync(bytes);
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

    public async Task LoadDocxAsync(byte[] docxBytes)
    {
        try
        {
            int timeout =0;
            while (!_isWebViewReady && timeout <50)
            {
                await Task.Delay(100);
                timeout++;
            }

            if (!_isWebViewReady)
                throw new InvalidOperationException("WebView is not ready yet.");

            string base64 = Convert.ToBase64String(docxBytes);
            string js = $"window.importDocxFromBase64('{EscapeJs(base64)}')";

            await Task.Delay(100);

            await _webView.EvaluateJavaScriptAsync(js);

            ImportCompleted?.Invoke(this, EventArgs.Empty);
        }
        catch (Exception ex)
        {
            ErrorOccurred?.Invoke(this, ex);
        }
    }

    public async Task<byte[]?> GetDocxAsync()
    {
        try
        {
            var js = "exportDocx()";
            var base64 = await _webView.EvaluateJavaScriptAsync(js);
            base64 = UnwrapJs(base64);
            if (string.IsNullOrWhiteSpace(base64)) return null;
            return Convert.FromBase64String(base64);
        }
        catch (Exception ex)
        {
            ErrorOccurred?.Invoke(this, ex);
            return null;
        }
    }

    static string EscapeJs(string s) => s.Replace("\\", "\\\\").Replace("'", "\\'").Replace("\r", "\\r").Replace("\n", "\\n");
    static string UnwrapJs(string s)
    {
        if (s.Length >=2 && s.StartsWith("\"") && s.EndsWith("\""))
            s = s.Substring(1, s.Length -2).Replace("\\n", "\n").Replace("\\r", "\r").Replace("\\\"", "\"");
        return s;
    }

    static string BuildHtml()
    {
        string GetRes(string name)
        {
            var assembly = typeof(WordFormView).Assembly;
            var resourceName = assembly.GetManifestResourceNames()
                .FirstOrDefault(r => r.EndsWith(name, StringComparison.OrdinalIgnoreCase));

            if (resourceName == null)
            {
                var available = string.Join(Environment.NewLine, assembly.GetManifestResourceNames());
                throw new InvalidOperationException($"Embedded resource '{name}' not found. Available:\n{available}");
            }

            using var stream = assembly.GetManifestResourceStream(resourceName)!;
            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }

        var quillCss = GetRes("quill.snow.css");
        var quillJs = GetRes("quill.min.js");
        var mammothJs = GetRes("mammoth.browser.min.js");
        var htmlDocxJs = GetRes("html-docx.js");

        var disableContextMenu = DeviceInfo.Current.Platform == DevicePlatform.MacCatalyst;
        var disableFlag = disableContextMenu ? "true" : "false";

        return $@"
<!doctype html>
<html>
<head>
<meta charset='utf-8'>
<style>{quillCss}</style>
<style>
html,body,#editor{{height:100%;margin:0;padding:0;background:white;}}
.ql-toolbar{{position:sticky;top:0;background:white;z-index:10;}}
</style>
</head>
<body>
<div id='editor'></div>
<script>{quillJs}</script>
<script>{mammothJs}</script>
<script>{htmlDocxJs}</script>
<script>
const DISABLE_CONTEXT_MENU = {disableFlag};
if (DISABLE_CONTEXT_MENU) {{
 window.addEventListener('contextmenu', function(e) {{ e.preventDefault(); }}, false);
}}

var quill;
(function(){{
 var toolbarOptions=[[{{'header':[1,2,3,false]}}],['bold','italic','underline'],['link','image'],[{{'list':'ordered'}},{{'list':'bullet'}}],['clean']];
 quill=new Quill('#editor',{{theme:'snow',modules:{{toolbar:toolbarOptions}}}});

 // Intercept right-click on images on Windows: provide a custom Save Image action
 if (!DISABLE_CONTEXT_MENU) {{
 document.addEventListener('contextmenu', function(e) {{
 const target = e.target;
 if (target && target.tagName === 'IMG') {{
 // Prevent default WebView2 menu, and trigger custom save path via postMessage
 e.preventDefault();
 try {{
 const canvas = document.createElement('canvas');
 canvas.width = target.naturalWidth || target.width;
 canvas.height = target.naturalHeight || target.height;
 const ctx = canvas.getContext('2d');
 ctx.drawImage(target,0,0);
 const dataUrl = canvas.toDataURL('image/png');
 const payload = JSON.stringify({{ type: 'saveImage', dataUrl }});
 window.chrome?.webview?.postMessage(payload);
 }} catch {{}}
 }}
 }}, false);
 }}
}})();

function blobToBase64(blob) {{
 return new Promise((res,rej) => {{
 var reader = new FileReader();
 reader.onloadend = () => res(reader.result.split(',')[1]);
 reader.onerror = rej;
 reader.readAsDataURL(blob);
 }});
}}

window.importDocxFromBase64 = async function(b64) {{
 try {{
 var bin = atob(b64);
 var bytes = new Uint8Array(bin.length);
 for (var i =0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
 var result = await mammoth.convertToHtml({{ arrayBuffer: bytes.buffer }});
 quill.root.innerHTML = result.value || '<p><em>No content</em></p>';
 return true;
 }} catch (e) {{ return false; }}
}};

window.getEditorHtml = function() {{ return quill.root.innerHTML; }};
window.setEditorHtml = function(html) {{ quill.root.innerHTML = html || ''; }};

window.exportDocx = async function() {{
 try {{
 var html = '<html><body>' + quill.root.innerHTML + '</body></html>';
 var blob = htmlDocx.asBlob(html);
 var base64 = await blobToBase64(blob);
 return base64;
 }} catch (e) {{ return ''; }}
}}
</script>
</body>
</html>";
    }
}