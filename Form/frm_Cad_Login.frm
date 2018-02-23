VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Cad_Login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alteração de Senha"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frm_Cad_Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtpss_liberacao_Conf 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2430
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2100
      Width           =   1320
   End
   Begin VB.TextBox txtpss_Login_Conf 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2430
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1110
      Width           =   1320
   End
   Begin VB.TextBox txtpss_liberacao 
      DataField       =   "pss_liberacao"
      DataMember      =   "Tab_Config"
      DataSource      =   "de"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2430
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1740
      Width           =   1320
   End
   Begin VB.TextBox txtpss_Login 
      DataField       =   "pss_Login"
      DataMember      =   "Tab_Config"
      DataSource      =   "de"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2430
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   765
      Width           =   1320
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fechar"
            Object.ToolTipText     =   "Fechar (Alt+F)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salvar"
            Object.ToolTipText     =   "Salvar Alteração (Alt+S)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar Alteração (Alt+C)"
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4680
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Login.frx":1CFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Login.frx":2014
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Login.frx":232E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Login.frx":2648
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Login.frx":2962
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Login.frx":2C7C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4080
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Liberação Confirmação:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   315
      TabIndex        =   8
      Top             =   2145
      Width           =   2025
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login Confirmação:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   315
      TabIndex        =   7
      Top             =   1155
      Width           =   1650
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Liberação:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   315
      TabIndex        =   6
      Top             =   1785
      Width           =   1155
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   315
      TabIndex        =   5
      Top             =   810
      Width           =   780
   End
End
Attribute VB_Name = "frm_Cad_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BarraF_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err1
   
    Select Case Button.key
        Case "fechar": Fechar
        Case "salvar": Salvar
        Case "cancelar": Cancelar
    End Select

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub




'*** Rotinas ***

Private Sub Fechar()
On Error GoTo err1
    Unload Me
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub





'--------- Ao Pressionar uma Tecla -----------
Private Sub txt_DESC_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_OP_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub GRID_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub




' -------  Teclas de Atalhos --------
Sub Keys(KeyCode As Integer, Shift As Integer)
    '*** Shift (4 = Alt) ***
    If Shift = 4 Then
        Select Case KeyCode
        Case 70: ' "F"
                Fechar
        Case 83: ' "S"
                Salvar
        Case 67: ' "C"
                Cancelar
        End Select
    End If
End Sub

Private Sub Salvar()
On Error GoTo err1
    
    
    If txtpss_Login = txtpss_Login_Conf And txtpss_liberacao = txtpss_liberacao_Conf Then
        
        de.rsTab_Config.UpdateBatch
        
    ElseIf txtpss_Login = txtpss_Login_Conf Then
        MsgBox "A senha de Login está incorreta!" & Chr(13) & "Redigite-as novamente!", vbCritical
        txtpss_Login = ""
        txtpss_Login_Conf = ""
        txtpss_Login.SetFocus
    
    ElseIf txtpss_liberacao = txtpss_liberacao_Conf Then
        MsgBox "A senha de Liberação está incorreta!" & Chr(13) & "Redigite-as novamente!", vbCritical
        txtpss_liberacao = ""
        txtpss_liberacao_Conf = ""
        txtpss_liberacao.SetFocus
    
    End If
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
    
End Sub


Private Sub Cancelar()

On Error GoTo err1
    
    de.rsTab_Config.CancelUpdate
    de.rsTab_Config.Requery
    
    txtpss_liberacao_Conf = txtpss_liberacao
    txtpss_Login_Conf = txtpss_Login
    txtpss_Login.SetFocus
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Form_Load()
    If de.rsTab_Config.EOF Then de.rsTab_Config.AddNew
End Sub

Private Sub txtpss_liberacao_Conf_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1) Then Salvar
     SendKeys "{tab}"
End If
End Sub

Private Sub txtpss_liberacao_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub txtpss_Login_Conf_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub txtpss_Login_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub
