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
                        End If
                        If Read Like "B2 = *" Then
                            B2 = Replace(Read, "B2 = ", "")
                        End If
                        If Read Like "B3 = *" Then
                            B3 = Replace(Read, "B3 = ", "")
                        End If
                        If Read Like "B4 = *" Then
                            B4 = Replace(Read, "B4 = ", "")
                        End If
                        If Read Like "B5 = *" Then
                            B5 = Replace(Read, "B5 = ", "")
                        End If
                        If Read Like "B6 = *" Then
                            B6 = Replace(Read, "B6 = ", "")
                        End If
                        If Read Like "StdAtt = *" Then
                            StdAtt = Replace(Read, "StdAtt = ", "")
                        End If
                        If Read Like "StdLink = *" Then
                            StdLink = Replace(Read, "StdLink = ", "")
                        End If
                Loop
            Close #ReadFile
Else
            
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set a = fs.CreateTextFile(SettingsFile, True, 0)
            a.WriteLine ("SelectedBrowser = 2")
            a.WriteLine ("B1 = http://direct.sussex.ac.uk")
            a.WriteLine ("B2 = http://direct.sussex.ac.uk")
            a.WriteLine ("B3 = http://direct.sussex.ac.uk")
            a.WriteLine ("B4 = http://direct.sussex.ac.uk")
            a.WriteLine ("B5 = http://direct.sussex.ac.uk")
            a.WriteLine ("B6 = http://direct.sussex.ac.uk")
            a.WriteLine ("StdAtt = 0")
            a.WriteLine ("StdLink = 0")
            a.Close
            MsgBox "The Settings File could not be found in " & SettingsFile & vbCrLf & "New Settings File Created.", vbCritical + vbOKOnly, "File not found"
End If
     

    
'Notify.NotifyMsg.Caption = "READ SETTINGS" & vbCrLf & _
        "SelBrow = " & SelectedBrowser & vbCrLf & _
        "B1 = " & B1 & vbCrLf & _
        "B2 = " & B2 & vbCrLf & _
        "B3 = " & B3 & vbCrLf & _
        "B4 = " & B4 & vbCrLf & _
        "B5 = " & B5 & vbCrLf & _
        "B6 = " & B6 & vbCrLf & _
        "StdAtt = " & StdAtt & vbCrLf & _
        "StdLink = " & StdLink
            
        'MPSAdminWidget.Enabled = False
        'MPSAdminSettings.Enabled = False
        'LinkForm.Enabled = False
        'Notify.Show vbModeless
        Exit Sub
     
Handler:
        MsgBox Err.Number & " " & Err.Description
        Exit Sub

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
    Print #WriteFile, "StdAtt = " & StdAtt
    Print #WriteFile, "StdLink = " & StdLink
Close #WriteFile

End Sub
