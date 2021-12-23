VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_MainMenu 
   Caption         =   "Menu de correos"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5610
   OleObjectBlob   =   "UF_MainMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ArchivadoCompras_Click()
    Call SuggestFolders(Array(), "username")
End Sub

Private Sub Bt_CargaCompleta_Click()
    Call MoveToFolder(Array("pstfoldername", "subfolder"))
End Sub

Private Sub Bt_Grupaje_Click()
    Call SuggestFolders(Array("pstfoldername", "subfolder1", "subfolder2"))
End Sub



