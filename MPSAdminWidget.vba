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
End Sub


' THESE WILL ALL NEED TO BE AMENDED TO ADAPT TO SETTINGS
Private Sub WB1_Click()
' i.e. Call MPSAdminCode.WB1
    Call MPSAdminCode.BIS
End Sub

Private Sub WB2_Click()
    Call MPSAdminCode.WCM
End Sub

Private Sub WB3_Click()
    Call MPSAdminCode.SxD
End Sub

Private Sub WB4_Click()
    Call MPSAdminCode.SyD
End Sub

Private Sub WB5_Click()
    Call MPSAdminCode.CMS
End Sub

Private Sub WB6_Click()
    Call MPSAdminCode.Cognos
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

