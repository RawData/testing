private void ThisAddIn_Startup(object sender, System.EventArgs e)
{
    this.Application.WorkbookActivate += new Excel.AppEvents_WorkbookActivateEventHandler(Application_WorkbookActivate);
}

void Application_WorkbookActivate(Excel.Workbook Wb)
{
    var commandBars = this.Application.CommandBars;
    var menuBar = commandBars["Worksheet Menu Bar"];
    var controls = menuBar.Controls;
    var commandBarControl = controls.Add(MsoControlType.msoControlButton, missing, missing, missing, true);
    var commandBarButton = (CommandBarButton)commandBarControl;
    commandBarButton.Style = MsoButtonStyle.msoButtonCaption;
    commandBarButton.Caption = "My Button";
    commandBarButton.Tag = "My Button";
    commandBarButton.Click += new _CommandBarButtonEvents_ClickEventHandler(CommandBarButton_Click);
}

void CommandBarButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
{
    MessageBox.Show("Button clicked.");
}
