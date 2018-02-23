VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Cad_Tp_Conta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CADASTRO DE TIPOS DE CONTAS"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   Icon            =   "frm_Cad_Tp_Conta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_codigo 
      DataField       =   "tp_cod"
      DataSource      =   "ADOREG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   7
      Top             =   1320
      Width           =   810
   End
   Begin VB.ComboBox TXT_NIVEL 
      DataField       =   "TP_NIVEL"
      DataSource      =   "ADOREG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frm_Cad_Tp_Conta.frx":1CFA
      Left            =   6240
      List            =   "frm_Cad_Tp_Conta.frx":1D0D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "P (PAGAMENTOS)         F - (FÉRIAS)        O - (OUTROS)"
      Top             =   1305
      Width           =   615
   End
   Begin VB.TextBox TXT_DESC 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   3975
   End
   Begin VB.ComboBox TXT_OP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frm_Cad_Tp_Conta.frx":1D20
      Left            =   5520
      List            =   "frm_Cad_Tp_Conta.frx":1D2D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1305
      Width           =   615
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   1429
      ButtonWidth     =   1482
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Fechar"
            Key             =   "fechar"
            Object.ToolTipText     =   "Fechar (Alt+F)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "salvar"
            Object.ToolTipText     =   "Salvar Alteração (Alt+S)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
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
               Picture         =   "frm_Cad_Tp_Conta.frx":1D3A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Tp_Conta.frx":2054
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Tp_Conta.frx":236E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Tp_Conta.frx":2688
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Tp_Conta.frx":29A2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Tp_Conta.frx":2CBC
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CÓDIGO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NÍVEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   6
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIÇÃO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "OP."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   3
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape txt_cod 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   120
      Top             =   840
      Width           =   6975
   End
End
Attribute VB_Name = "frm_Cad_Tp_Conta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ultimoCod As Integer

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
Private Sub Cancelar()
On Error GoTo err1
    TXT_DESC = ""
    TXT_OP = "+"
    TXT_NIVEL = "0"
    TXT_DESC.SetFocus

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Fechar()
On Error GoTo err1
    If de.rsTAB_TP_CONTA.State = 1 Then de.rsTAB_TP_CONTA.Requery
    Unload Me
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Salvar()
On Error GoTo err1
    If TXT_DESC <> "" And TXT_OP <> "" And txt_codigo <> "" Then
        de.cmdIncluirTpConta txt_codigo, TXT_DESC, TXT_OP, TXT_NIVEL
        de.cmdIncluirLog Date, Time, w_usuario, "INCLUIR", "TIPO DE CONTA", "CÓDIGO: " & txt_codigo & " | DESCRIÇÃO: " & TXT_DESC & " | OP: " & TXT_OP & " | NÍVEL: " & TXT_NIVEL
        MsgBox "Registro salvo com sucesso!", vbInformation
        txt_codigo = txt_codigo + 1
        Cancelar
    Else
        MsgBox "Preencha os Campos!", vbCritical
    End If
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
    
End Sub







Private Sub Text1_Change()

End Sub

Private Sub Form_Load()
    ultimoCod = de.cnc.Execute("SELECT MAX(TP_COD) FROM TAB_TP_CONTA").Fields(0)
    txt_codigo.Text = ultimoCod + 1
End Sub

Private Sub TXT_DESC_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub TXT_DESC_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub

'--------- Ao Pressionar uma Tecla -----------
Private Sub txt_DESC_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub TXT_NIVEL_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1) Then Salvar
    SendKeys "{tab}"
 End If
End Sub

Private Sub TXT_OP_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then SendKeys "{tab}"
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

