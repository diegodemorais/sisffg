VERSION 5.00
Begin VB.Form frm_Acesso_Especial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acesso Especial"
   ClientHeight    =   600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   2790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtSenha 
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frm_Acesso_Especial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOk_Click()
    
    w_pass = de.cnc.Execute("SELECT pss_admin FROM  Tab_Config;").Fields(0)
    
    If txtSenha.Text = w_pass Then
        w_AcessoEspecial = True
    Else
        w_AcessoEspecial = False
    End If
        
    Me.Hide
    
    'MsgBox ("Senha de Acesso Especial definida com sucesso!")
    
End Sub


Private Sub txtSenha_Change()
    If KeyCode = 13 Then
        btnOk_Click
    End If
End Sub
