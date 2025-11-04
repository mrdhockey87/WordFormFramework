using Microsoft.Maui.Controls;
using WordFormFramework.Controls;

namespace WordFormFramework.Demo;

public partial class EditorPage : ContentPage
{
    private readonly string _path;

    public EditorPage(string path)
    {
        InitializeComponent();
        _path = path;
        Appearing += OnAppearingPage;
    }

    private async void OnAppearingPage(object sender, EventArgs e)
    {
        await wordView.OpenDocxFileAsync(_path);
    }

    private async void OnSaveClicked(object sender, EventArgs e)
    {
        var savePath = Path.Combine(FileSystem.CacheDirectory, "Saved.docx");
        await wordView.SaveDocxToFileAsync(savePath);
        await DisplayAlert("Saved", $"Saved to {savePath}", "OK");
    }

    private async void OnBackClicked(object sender, EventArgs e)
    {
        await Navigation.PopAsync();
    }
}
