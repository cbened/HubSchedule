namespace HubSchedule.OutlookAddIn;

internal static class RibbonXml
{
    public const string Definition = @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='Ribbon_Load'>
  <ribbon>
    <tabs>
      <tab idMso='TabReadMessage'>
        <group id='HubScheduleGroup.Read' label='HubSchedule'>
          <toggleButton id='HubScheduleTogglePane' label='Open Submissions' size='large' onAction='OnTogglePane' getPressed='GetPressed'/>
        </group>
      </tab>
      <tab idMso='TabAppointment'>
        <group id='HubScheduleGroup.Calendar' label='HubSchedule'>
          <toggleButton id='HubScheduleTogglePaneAppointment' label='Open Submissions' size='large' onAction='OnTogglePane' getPressed='GetPressed'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
}
