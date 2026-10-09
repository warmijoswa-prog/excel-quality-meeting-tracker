VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Usuario"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   1965
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
usuarioc = UserForm3.TextBox1.Value
Unload UserForm3
End Sub

Private Sub CommandButton1_Enter()
usuarioc = UserForm3.TextBox1.Value
Unload UserForm3
End Sub



Private Sub TextBox1_Change()
TextBox1 = UCase(TextBox1)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then salir3 = 10
End Sub
