VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Nuevo Registro"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
noc = True
End Sub

Private Sub CommandButton1_Click()
eebb = UserForm2.TextBox1.Value
servicios = UserForm2.TextBox2.Value
ciudad = UserForm2.TextBox3.Value
escalaanoc = UserForm2.TextBox4.Value
Unload UserForm2
End Sub


Private Sub TextBox1_Change()
TextBox1 = UCase(TextBox1)
End Sub


Private Sub TextBox2_Change()
TextBox2 = UCase(TextBox2)
End Sub

Private Sub TextBox4_Change()
TextBox4 = UCase(TextBox4)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then salir2 = 10
End Sub
