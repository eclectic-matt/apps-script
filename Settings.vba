Public a As Object
Public fs As Object

' Read the settings in BEFORE the Widget opens up (fired on startup)
' This creates a settings file if none exists so no FileNotFound errors
Public Sub ReadSettings()

'Start error handling
On Error GoTo Handler:

    'Declaring variables
    TempDir = Environ("Temp")
    SettingsFileName = "\MPSAdminWidgetSettings.txt"
    SettingsFile = TempDir & SettingsFileName
    ' FreeFile? huh? Necessary apparently
    ReadFile = FreeFile()

If Dir(SettingsFile) <> "" Then
            'Read all settings from file
            Open SettingsFile For Input As #ReadFile
                Do While Not EOF(ReadFile)
                    Line Input #ReadFile, Read
                        If Read Like "SelectedBrowser = *" Then
                            SelectedBrowser = Replace(Read, "SelectedBrowser = ", "")
                        End If
                        If Read Like "B1 = *" Then
                            B1 = Replace(Read, "B1 = ", "")
                            myLinks(1, 0) = B1
                        End If
                        If Read Like "B1F = *" Then
                            B1F = Replace(Read, "B1F = ", "")
                            myLinks(0, 0) = B1F
                        End If
                        If Read Like "B2 = *" Then
                            B2 = Replace(Read, "B2 = ", "")
                            myLinks(1, 1) = B2
                        End If
                        If Read Like "B2F = *" Then
                            B2F = Replace(Read, "B2F = ", "")
                            myLinks(0, 1) = B2F
                        End If
                        If Read Like "B3 = *" Then
                            B3 = Replace(Read, "B3 = ", "")
                            myLinks(1, 2) = B3
                        End If
                        If Read Like "B3F = *" Then
                            B3F = Replace(Read, "B3F = ", "")
                            myLinks(0, 2) = B3F
                        End If
                        If Read Like "B4 = *" Then
                            B4 = Replace(Read, "B4 = ", "")
                            myLinks(1, 3) = B4
                        End If
                        If Read Like "B4F = *" Then
                            B4F = Replace(Read, "B4F = ", "")
                            myLinks(0, 3) = B4F
                        End If
                        If Read Like "B5 = *" Then
                            B5 = Replace(Read, "B5 = ", "")
                            myLinks(1, 4) = B5
                        End If
                        If Read Like "B5F = *" Then
                            B5F = Replace(Read, "B5F = ", "")
                            myLinks(0, 4) = B5F
                        End If
                        If Read Like "B6 = *" Then
                            B6 = Replace(Read, "B6 = ", "")
                            myLinks(1, 5) = B6
                        End If
                        If Read Like "B6F = *" Then
                            B6F = Replace(Read, "B6F = ", "")
                            myLinks(0, 5) = B6F
                        End If
                        If Read Like "StdAtt = *" Then
                            StdAtt = Replace(Read, "StdAtt = ", "")
                        End If
                        If Read Like "StdLink = *" Then
                            StdLink = Replace(Read, "StdLink = ", "")
                        End If
                Loop
            Close #ReadFile
                Dim MyStr As String
                MyStr = "Links 1 - 6:" & vbCrLf
                For i = 0 To 5
                    MyStr = MyStr & myLinks(0, i) & vbCrLf
                Next
                MsgBox (MyStr)
Else
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set a = fs.CreateTextFile(SettingsFile, True, 0)
            a.WriteLine ("SelectedBrowser = 2")
            a.WriteLine ("B1 = http://direct.sussex.ac.uk")
            a.WriteLine ("B2 = https://studydirect.sussex.ac.uk/login/")
            a.WriteLine ("B3 = http://www.sussex.ac.uk/its/services/staffservices/businessapplications/")
            a.WriteLine ("B4 = http://w2k3bis4.admin.sussex.ac.uk/forms/frmservlet?config=bisapps-java")
            a.WriteLine ("B5 = http://www.sussex.ac.uk/wcm/entry/")
            a.WriteLine ("B6 = https://abw.admin.sussex.ac.uk/Agresso/System/Login.aspx")
            a.WriteLine ("B1F = Sussex Direct")
            a.WriteLine ("B2F = Study Direct")
            a.WriteLine ("B3F = BIS Logon")
            a.WriteLine ("B4F = CMS")
            a.WriteLine ("B5F = WCM")
            a.WriteLine ("B6F = Agresso")
            a.WriteLine ("StdAtt = 0")
            a.WriteLine ("StdLink = 0")
            a.Close
            B1 = "http://direct.sussex.ac.uk"
            B2 = "https://studydirect.sussex.ac.uk/login/"
            B3 = "http://www.sussex.ac.uk/its/services/staffservices/businessapplications/"
            B4 = "http://w2k3bis4.admin.sussex.ac.uk/forms/frmservlet?config=bisapps-java"
            B5 = "http://www.sussex.ac.uk/wcm/entry/"
            B6 = "https://abw.admin.sussex.ac.uk/Agresso/System/Login.aspx"
            B1F = "Sussex Direct"
            B2F = "Study Direct"
            B3F = "BIS Logon"
            B4F = "CMS"
            B5F = "WCM"
            B6F = "Agresso"
            myLinks(1, 0) = B1
            myLinks(1, 1) = B2
            myLinks(1, 2) = B3
            myLinks(1, 3) = B4
            myLinks(1, 4) = B5
            myLinks(1, 5) = B6
            myLinks(0, 0) = B1F
            myLinks(0, 1) = B2F
            myLinks(0, 2) = B3F
            myLinks(0, 3) = B4F
            myLinks(0, 4) = B5F
            myLinks(0, 5) = B6F
            MsgBox "The Settings File could not be found in " & SettingsFile & vbCrLf & "New Settings File Created.", vbCritical + vbOKOnly, "File not found"
End If

Call Captions

        Exit Sub
     
Handler:
        MsgBox Err.Number & " " & Err.Description
        Exit Sub

End Sub

Public Sub Captions()
MPSAdminWidget.WB1.Caption = B1F
MPSAdminWidget.WB2.Caption = B2F
MPSAdminWidget.WB3.Caption = B3F
MPSAdminWidget.WB4.Caption = B4F
MPSAdminWidget.WB5.Caption = B5F
MPSAdminWidget.WB6.Caption = B6F
MPSAdminSettings.B1edit.Caption = B1F
MPSAdminSettings.B2edit.Caption = B2F
MPSAdminSettings.B3edit.Caption = B3F
MPSAdminSettings.B4edit.Caption = B4F
MPSAdminSettings.B5edit.Caption = B5F
MPSAdminSettings.B6edit.Caption = B6F
End Sub

Public Sub SaveSettings()

'Declaring variables
TempDir = Environ("Temp")
SettingsFileName = "\MPSAdminWidgetSettings.txt"
SettingsFile = TempDir & SettingsFileName
WriteFile = FreeFile()

'Write all settings to file
Open SettingsFile For Output As #WriteFile
    Print #WriteFile, "SelectedBrowser = " & SelectedBrowser
    Print #WriteFile, "B1 = " & B1
    Print #WriteFile, "B2 = " & B2
    Print #WriteFile, "B3 = " & B3
    Print #WriteFile, "B4 = " & B4
    Print #WriteFile, "B5 = " & B5
    Print #WriteFile, "B6 = " & B6
    Print #WriteFile, "B1F = " & B1F
    Print #WriteFile, "B2F = " & B2F
    Print #WriteFile, "B3F = " & B3F
    Print #WriteFile, "B4F = " & B4F
    Print #WriteFile, "B5F = " & B5F
    Print #WriteFile, "B6F = " & B6F
    Print #WriteFile, "StdAtt = " & StdAtt
    Print #WriteFile, "StdLink = " & StdLink
Close #WriteFile

End Sub
