VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRun_Click()
    Dim lastName As String
    Dim i As Integer
    Dim char As String
    Dim isValid As Boolean
    
    lastName = txtLastName.Text
    
    If Len(lastName) = 0 Then
        MsgBox "Please enter a last name.", vbExclamation, "Input Required"
        txtLastName.SetFocus
        Exit Sub
    End If
    
    isValid = True
    
    For i = 1 To Len(lastName)
        char = Mid(lastName, i, 1)
        
        If Not ((char >= "A" And char <= "Z") Or _
                (char >= "a" And char <= "z") Or _
                char = " " Or char = "-" Or char = "'") Then
            isValid = False
            Exit For
        End If
    Next i
    
    If isValid Then
        MsgBox "Valid last name entered: " & lastName, vbInformation, "Success"
    Else
        MsgBox "Invalid characters detected! Last name should contain only alphabetical characters, spaces, hyphens, or apostrophes.", _
               vbExclamation, "Invalid Input"
        txtLastName.Text = ""
        txtLastName.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
