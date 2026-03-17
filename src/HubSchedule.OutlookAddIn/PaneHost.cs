using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace HubSchedule.OutlookAddIn;

internal sealed class PaneHost : IDisposable
{
    private readonly HubScheduleWindow _window;

    public PaneHost(Outlook.Inspector inspector)
    {
        _window = new HubScheduleWindow();
        _window.Text = $"HubSchedule - {inspector.Caption}";
    }

    public bool Visible
    {
        get => _window.Visible;
        set
        {
            if (value)
            {
                _window.Show();
                _window.BringToFront();
            }
            else
            {
                _window.Hide();
            }
        }
    }

    public void Dispose()
    {
        _window.Dispose();
    }
}
