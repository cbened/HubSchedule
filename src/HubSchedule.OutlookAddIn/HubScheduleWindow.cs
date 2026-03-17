using System;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;

namespace HubSchedule.OutlookAddIn;

public class HubScheduleWindow : Form
{
    private const string HubSpotSubmissionsUrl = "https://app.hubspot.com/submissions";

    public HubScheduleWindow()
    {
        Width = 480;
        Height = 720;
        StartPosition = FormStartPosition.CenterScreen;

        var browser = new WebView2
        {
            Dock = DockStyle.Fill
        };

        Controls.Add(browser);

        Load += async (_, _) =>
        {
            try
            {
                await browser.EnsureCoreWebView2Async();
                browser.CoreWebView2.Navigate(HubSpotSubmissionsUrl);
            }
            catch (Exception ex)
            {
                Controls.Clear();
                Controls.Add(new Label
                {
                    Dock = DockStyle.Fill,
                    Text = "Unable to initialize embedded browser. Install Microsoft Edge WebView2 Runtime.\n\n" + ex.Message,
                    Padding = new Padding(12)
                });
            }
        };
    }
}
