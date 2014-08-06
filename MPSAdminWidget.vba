' ADMIN WIDGET ACTIVATE
Private Sub UserForm_Activate()
    Me.StartUpPosition = START_POS
End Sub

Private Sub Gbutton_Click()
Call Shell("explorer.exe" & " " & "G:\", vbNormalFocus)
End Sub

Private Sub Nbutton_Click()
Call Shell("explorer.exe" & " " & "N:\", vbNormalFocus)
End Sub

Private Sub UserForm_Initialize()
    MPSAdminWidget.StartUpPosition = START_POS
    Call Settings.Captions
End Sub

Private Sub StaffSearchBox_AfterUpdate()
    Call MPSAdminCode.StaffSearch
End Sub

Private Sub SussexSearchBox_AfterUpdate()
    Call MPSAdminCode.SussexSearch
End Sub

Private Sub StudentSearchBox_AfterUpdate()
    Call MPSAdminCode.StudentSearch
End Sub

Private Sub StaffSearch_Click()
    Call MPSAdminCode.StaffSearch
End Sub

Private Sub StudentSearch_Click()
    Call MPSAdminCode.StudentSearch
End Sub

Private Sub SussexSearch_Click()
    Call MPSAdminCode.SussexSearch
End Sub

Private Sub WB1_Click()
    Call MPSAdminCode.WB1
End Sub

Private Sub WB2_Click()
    Call MPSAdminCode.WB2
End Sub

Private Sub WB3_Click()
    Call MPSAdminCode.WB3
End Sub

Private Sub WB4_Click()
    Call MPSAdminCode.WB4
End Sub

Private Sub WB5_Click()
    Call MPSAdminCode.WB5
End Sub

Private Sub WB6_Click()
    Call MPSAdminCode.WB6
End Sub

Private Sub RefreshButton_Click()
    Settings.ReadSettings
    MPSAdminWidget.StaffSearchBox.Text = ""
    MPSAdminWidget.StudentSearchBox.Text = ""
    MPSAdminWidget.SussexSearchBox.Text = ""
    MPSAdminCode.WidgetShow
End Sub

Private Sub WidgetSettingsBtn_Click()
    MPSAdminWidget.Hide
    Load MPSAdminSettings
    MPSAdminSettings.Show vbModeless
End Sub
