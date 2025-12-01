VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Scan My Kad "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim tData As MyKadData
    Dim sMsg As String

    Screen.MousePointer = vbHourglass

    If modMyKad.PerformMyKadScan(tData) Then
        Screen.MousePointer = vbNormal

        ' --- PERSONAL INFO ---
        sMsg = "==============================" & vbCrLf
        sMsg = sMsg & "PERSONAL INFO" & vbCrLf
        sMsg = sMsg & "==============================" & vbCrLf
        sMsg = sMsg & "Name            : " & tData.Name & vbCrLf
        sMsg = sMsg & "IC No           : " & tData.ICNo & vbCrLf
        sMsg = sMsg & "Old IC          : " & tData.OldIC & vbCrLf
        sMsg = sMsg & "Gender          : " & tData.Gender & vbCrLf
        sMsg = sMsg & "Date of Birth   : " & tData.DOB & vbCrLf
        sMsg = sMsg & "Place of Birth  : " & tData.PlaceOfBirth & vbCrLf
        sMsg = sMsg & "Race            : " & tData.Race & vbCrLf
        sMsg = sMsg & "Religion        : " & tData.Religion & vbCrLf
        sMsg = sMsg & "Citizenship     : " & tData.Citizenship & vbCrLf & vbCrLf

        ' --- ADDRESS INFO ---
        sMsg = sMsg & "==============================" & vbCrLf
        sMsg = sMsg & "ADDRESS INFO" & vbCrLf
        sMsg = sMsg & "==============================" & vbCrLf
        sMsg = sMsg & tData.Address1 & vbCrLf
        sMsg = sMsg & tData.Address2 & vbCrLf
        sMsg = sMsg & tData.Address3 & vbCrLf
        sMsg = sMsg & tData.Postcode & " " & tData.City & vbCrLf
        sMsg = sMsg & tData.State & vbCrLf
        sMsg = sMsg & "==============================" & vbCrLf

        ' Display Results
        MsgBox sMsg, vbInformation, "MyKad Details"

    Else
        Screen.MousePointer = vbNormal
        MsgBox "Failed to read MyKad." & vbCrLf & _
               "Please check reader connection.", vbCritical
    End If
End Sub


