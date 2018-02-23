VERSION 5.00
Begin VB.Form frm_Alt_Acesso_Especial 
   Caption         =   "Alteração de Acesso Especial"
   ClientHeight    =   975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtSenha 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Definição de SENHA para Acesso Especial:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frm_Alt_Acesso_Especial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'= Global Variables
Dim md5Test As MD5



Private Sub Form_Load()
    ' Instantiate our class
    Set md5Test = New MD5
End Sub


Private Sub btnOk_Click()
    
    If txtSenha.Text = "" Then
        MsgBox ("A senha não pode ser em branca. Digite alguma senha.")
    Else
        de.cnc.Execute "Update Tab_Config Set pss_admin = '" & md5Test.DigestStrToHexStr(LCase(txtSenha.Text)) & "';"
        de.cnc.Execute "Update Tab_Config Set pss_liberacao = '" & md5Test.DigestStrToHexStr(LCase(txtSenha.Text)) & "';"
        de.cnc.Execute "Update Tab_Config Set pss_login = '" & md5Test.DigestStrToHexStr(LCase(txtSenha.Text)) & "';"
        MsgBox ("Senha de Acesso Especial definida com sucesso!")
    End If
    
End Sub

Private Sub txtSenha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        btnOk_Click
    End If
End Sub
