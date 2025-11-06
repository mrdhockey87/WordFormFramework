using Microsoft.Maui.Controls;
using WordFormFramework.Controls;

namespace WordFormFramework.Demo;

public partial class MainPage : ContentPage
{
    public MainPage()
    {
        InitializeComponent();
    }

    private async void OnOpenClicked(object sender, EventArgs e)
    {
        try
        {
            var docxOnly = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
            {
                // iOS/Mac: UTType for .docx
                { DevicePlatform.iOS, new[] { "org.openxmlformats.wordprocessingml.document" } },
                { DevicePlatform.MacCatalyst, new[] { "org.openxmlformats.wordprocessingml.document" } },

                // Android: MIME type for .docx
                { DevicePlatform.Android, new[] { "application/vnd.openxmlformats-officedocument.wordprocessingml.document" } },

                // Windows: extension for .docx
                { DevicePlatform.WinUI, new[] { ".docx" } },
            });

            var result = await FilePicker.Default.PickAsync(new PickOptions
            {
                PickerTitle = "Select a Word (.docx) document",
                FileTypes = docxOnly
            });

            if (result == null) return;

            // Defensive check (in case a platform returns other types)
            if (!result.FileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            {
                await DisplayAlert("Unsupported file", "Please select a .docx file.", "OK");
                return;
            }

            await Navigation.PushAsync(new EditorPage(result.FullPath, isRtf:false));
        }
        catch (Exception ex)
        {
            await DisplayAlert("Error", ex.Message, "OK");
        }
    }

    private async void OnOpenRtfPromptClicked(object sender, EventArgs e)
    {
        try
        {
            var rtfType = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
            {
                { DevicePlatform.iOS, new[] { "public.rtf" } },
                { DevicePlatform.MacCatalyst, new[] { "public.rtf" } },
                { DevicePlatform.Android, new[] { "application/rtf", "text/rtf" } },
                { DevicePlatform.WinUI, new[] { ".rtf" } },
            });
            var result = await FilePicker.Default.PickAsync(new PickOptions
            {
                PickerTitle = "Select an RTF document (Prompt Lock Toggle)",
                FileTypes = rtfType
            });
            if (result == null) return;
            if (!result.FileName.EndsWith(".rtf", StringComparison.OrdinalIgnoreCase))
            {
                await DisplayAlert("Unsupported file", "Please select a .rtf file.", "OK");
                return;
            }
            await Navigation.PushAsync(new EditorPage(result.FullPath, isRtf:true, promptLockToggle:true));
        }
        catch (Exception ex)
        {
            await DisplayAlert("Error", ex.Message, "OK");
        }
    }

    private async void OnOpenRtfLockedClicked(object sender, EventArgs e)
    {
        try
        {
            var rtfType = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
            {
                { DevicePlatform.iOS, new[] { "public.rtf" } },
                { DevicePlatform.MacCatalyst, new[] { "public.rtf" } },
                { DevicePlatform.Android, new[] { "application/rtf", "text/rtf" } },
                { DevicePlatform.WinUI, new[] { ".rtf" } },
            });
            var result = await FilePicker.Default.PickAsync(new PickOptions
            {
                PickerTitle = "Select an RTF document (Locked - No Toggle)",
                FileTypes = rtfType
            });
            if (result == null) return;
            if (!result.FileName.EndsWith(".rtf", StringComparison.OrdinalIgnoreCase))
            {
                await DisplayAlert("Unsupported file", "Please select a .rtf file.", "OK");
                return;
            }
            await Navigation.PushAsync(new EditorPage(result.FullPath, isRtf:true, promptLockToggle:false));
        }
        catch (Exception ex)
        {
            await DisplayAlert("Error", ex.Message, "OK");
        }
    }
}
