Private Sub StaffSearch_Click()
    Call MPSAdminCode.StaffSearch
End Sub

Private Sub StudentSearch_Click()
    Call MPSAdminCode.StudentSearch
End Sub

' ADMIN WIDGET ACTIVATE
Private Sub UserForm_Activate()
    Me.StartUpPosition = START_POS
End Sub

Private Sub UserForm_Initialize()
    MPSAdminWidget.StartUpPosition = START_POS
    Call Settings.Captions
End Sub

' THESE WILL ALL NEED TO BE AMENDED TO ADAPT TO SETTINGS
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

Private Sub StaffSearchBox_AfterUpdate()
    Call MPSAdminCode.StaffSearch
End Sub

Private Sub StudentSearchBox_AfterUpdate()
    Call MPSAdminCode.StudentSearch
End Sub

Private Sub WidgetSettingsBtn_Click()
    MPSAdminWidget.Hide
    Load MPSAdminSettings
    MPSAdminSettings.Show vbModeless
End Sub

