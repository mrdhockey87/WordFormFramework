using Microsoft.Maui.Controls;
using WordFormFramework.Controls;

namespace WordFormFramework.Demo;

public partial class EditorPage : ContentPage
{
    private readonly string _path;
    private readonly bool _isRtf;
    private readonly bool _promptLockToggle;

    public EditorPage(string path, bool isRtf, bool promptLockToggle = false)
    {
        InitializeComponent();
        _path = path;
        _isRtf = isRtf;
        _promptLockToggle = promptLockToggle;
        Appearing += OnAppearingPage;
    }

    private async void OnAppearingPage(object sender, EventArgs e)
    {
        if (_isRtf)
        {
            await wordView.OpenRtfFileAsync(_path);
        }
        else
        {
            await wordView.OpenDocxFileAsync(_path);
        }
    }

    private async void OnSaveDocxClicked(object sender, EventArgs e)
    {
        bool ok = await wordView.SaveDocxWithPickerAsync();
        await DisplayAlert(ok ? "Saved" : "Save Failed", ok ? "Document saved." : "Could not save file", "OK");
    }

    private async void OnSaveRtfClicked(object sender, EventArgs e)
    {
        bool ok = await wordView.SaveRtfWithPickerAsync();
        await DisplayAlert(ok ? "Saved" : "Save Failed", ok ? "Document saved." : "Could not save file", "OK");
    }

    private async void OnBackClicked(object sender, EventArgs e)
    {
        await Navigation.PopAsync();
    }

    private async void OnOpenRtfNoPrompt(object sender, EventArgs e)
    {
        if (!_isRtf) return;
        await wordView.OpenRtfFileAsync(_path);
    }

    private async void OnOpenRtfPrompt(object sender, EventArgs e)
    {
        if (!_isRtf) return;
        await wordView.OpenRtfFileAsync(_path);
    }
}
