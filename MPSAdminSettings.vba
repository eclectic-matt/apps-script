''' USERFORM ACTIVATE AND INITIALIZE CODE HERE '''
Private Sub UserForm_Initialize()
    MPSAdminSettings.StartUpPosition = START_POS
    MPSAdminSettings.BrowserSelectBox.AddItem "Firefox"
    MPSAdminSettings.BrowserSelectBox.AddItem "Chrome"
    MPSAdminSettings.BrowserSelectBox.AddItem "Internet Explorer"
End Sub

Private Sub UserForm_Activate()
    Me.StartUpPosition = START_POS
    Call Settings.Captions
End Sub

Public Sub B1edit_Click()
    SelButton = 1
    Call LinkBtnClicked
End Sub

Public Sub B2edit_Click()
    SelButton = 2
    Call LinkBtnClicked
End Sub

Public Sub B3edit_Click()
    SelButton = 3
    Call LinkBtnClicked
End Sub

Public Sub B4edit_Click()
    SelButton = 4
    Call LinkBtnClicked
End Sub

Public Sub B5edit_Click()
    SelButton = 5
    Call LinkBtnClicked
End Sub

Public Sub B6edit_Click()
    SelButton = 6
    Call LinkBtnClicked
End Sub

Public Sub LinkBtnClicked()
    MPSAdminSettings.Hide
    'Load LinkForm
    LinkForm.Show
End Sub

Public Sub BrowserSelectBox_AfterUpdate()

    Notify.NotifyMsg.Caption = "Your new browser is " & BrowserSelectBox.Value
    MPSAdminWidget.Enabled = False
    MPSAdminSettings.Enabled = False
    LinkForm.Enabled = False
    Notify.Show vbModeless
    ' 1 IE, 2 Firefox, 3 Chrome, 4 Safari.... maybe?
    Select Case BrowserSelectBox.Value
        Case "Internet Explorer"
            SelectedBrowser = 1
        Case "Firefox"
            SelectedBrowser = 2
        Case "Chrome"
            SelectedBrowser = 3
        Case Else
            SelectedBrowser = DEF_BROWSER
    End Select
        
End Sub

Private Sub CloseSettings_Click()
    MPSAdminSettings.Hide
    MPSAdminWidget.Show vbModeless
End Sub

Private Sub UserForm_Terminate()
    MPSAdminWidget.Show vbModeless
End Sub
