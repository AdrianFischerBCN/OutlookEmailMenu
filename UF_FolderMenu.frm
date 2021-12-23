VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_FolderMenu 
   Caption         =   "Propuesta de Carpetas"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7020
   OleObjectBlob   =   "UF_FolderMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_FolderMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub GetLabels_and_MoveEmail(ByVal Label As String)
    
    lenArray = UBound(UnivFolderList)
    ReDim Preserve UnivFolderList(0 To lenArray + 1)
    UnivFolderList(lenArray + 1) = Label
    Call MoveToFolder(UnivFolderList, UnivUserName)
    
    Unload UF_FolderMenu
    
End Sub


Private Sub Label1_Click()

    Call GetLabels_and_MoveEmail(Me.Label1.Caption)
End Sub

Private Sub Label2_Click()
    Call GetLabels_and_MoveEmail(Me.Label2.Caption)
End Sub

Private Sub Label3_Click()
    Call GetLabels_and_MoveEmail(Me.Label3.Caption)
End Sub

Private Sub Label4_Click()
    Call GetLabels_and_MoveEmail(Me.Label4.Caption)
End Sub

Private Sub Label5_Click()
    Call GetLabels_and_MoveEmail(Me.Label5.Caption)
End Sub

Private Sub Label6_Click()
    Call GetLabels_and_MoveEmail(Me.Label6.Caption)
End Sub

Private Sub Label7_Click()
    Call GetLabels_and_MoveEmail(Me.Label7.Caption)
End Sub

Private Sub Label8_Click()
    Call GetLabels_and_MoveEmail(Me.Label8.Caption)
End Sub




Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub

Sub ChangeCaptionName(Label As Integer)
    Dim s As String
    s = "Label" & Label
    
End Sub
