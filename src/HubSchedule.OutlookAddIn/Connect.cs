using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace HubSchedule.OutlookAddIn;

[ComVisible(true)]
[Guid("4FA2D5C9-0C7B-4CE0-A13F-A91FF0B2D8F4")]
[ProgId("HubSchedule.OutlookAddIn")]
public class Connect : IDTExtensibility2, IRibbonExtensibility
{
    private Outlook.Application? _application;
    private readonly Dictionary<Outlook.Inspector, PaneHost> _paneByInspector = new();
    private IRibbonUI? _ribbon;

    public string GetCustomUI(string ribbonId)
    {
        return RibbonXml.Definition;
    }

    public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
    {
        _application = (Outlook.Application)application;
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
    {
        foreach (var pane in _paneByInspector.Values)
        {
            pane.Dispose();
        }

        _paneByInspector.Clear();
        _application = null;
    }

    public void OnAddInsUpdate(ref Array custom)
    {
    }

    public void OnStartupComplete(ref Array custom)
    {
    }

    public void OnBeginShutdown(ref Array custom)
    {
    }

    public void Ribbon_Load(IRibbonUI ribbonUi)
    {
        _ribbon = ribbonUi;
    }

    public void OnTogglePane(IRibbonControl control, bool isPressed)
    {
        if (control.Context is not Outlook.Inspector inspector)
        {
            return;
        }

        var pane = GetOrCreatePane(inspector);
        pane.Visible = isPressed;
    }

    public bool GetPressed(IRibbonControl control)
    {
        if (control.Context is not Outlook.Inspector inspector)
        {
            return false;
        }

        return _paneByInspector.TryGetValue(inspector, out var pane) && pane.Visible;
    }

    private PaneHost GetOrCreatePane(Outlook.Inspector inspector)
    {
        if (_paneByInspector.TryGetValue(inspector, out var existing))
        {
            return existing;
        }

        var pane = new PaneHost(inspector);
        _paneByInspector[inspector] = pane;

        ((Outlook.InspectorEvents_10_Event)inspector).Close += () =>
        {
            if (_paneByInspector.Remove(inspector, out var closingPane))
            {
                closingPane.Dispose();
                _ribbon?.InvalidateControl("HubScheduleTogglePane");
                _ribbon?.InvalidateControl("HubScheduleTogglePaneAppointment");
            }
        };

        return pane;
    }
}
