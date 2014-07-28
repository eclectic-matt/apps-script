Private Sub NotifyOK_Click()
MPSAdminWidget.Enabled = True
MPSAdminSettings.Enabled = True
LinkForm.Enabled = True
Notify.Hide
End Sub

Private Sub UserForm_Initialize()
Notify.StartUpPosition = START_POS
End Sub
