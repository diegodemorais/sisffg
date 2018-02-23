VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_Cad_Desc_Calc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CADASTRO DE DESCRIÇÃO DE CÁLCULOS"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "frm_Cad_Desc_Calc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_valor 
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
      Left            =   3960
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox TXT_OP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "frm_Cad_Desc_Calc.frx":1CFA
      Left            =   4440
      List            =   "frm_Cad_Desc_Calc.frx":1D04
      TabIndex        =   5
      Text            =   "+"
      Top             =   3120
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
      Height          =   915
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   3855
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5430
      _ExtentX        =   9578
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
               Picture         =   "frm_Cad_Desc_Calc.frx":1D0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Desc_Calc.frx":2028
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Desc_Calc.frx":2342
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Desc_Calc.frx":265C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Desc_Calc.frx":2976
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Desc_Calc.frx":2C90
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSDataListLib.DataCombo TXT_FUNC 
      Bindings        =   "frm_Cad_Desc_Calc.frx":2FAA
      Height          =   360
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "F_NOME"
      BoundColumn     =   "M_F_COD"
      Text            =   ""
      Object.DataMember      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoFunc 
      Height          =   330
      Left            =   240
      Top             =   3600
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo TXT_CONTA 
      Bindings        =   "frm_Cad_Desc_Calc.frx":2FC0
      Height          =   360
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "TP_DESC"
      BoundColumn     =   "TP_COD"
      Text            =   ""
      Object.DataMember      =   "TAB_TP_CONTA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo TXT_OP_CONTA 
      Bindings        =   "frm_Cad_Desc_Calc.frx":2FD1
      Height          =   360
      Left            =   4440
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "TP_OP"
      BoundColumn     =   "TP_COD"
      Text            =   ""
      Object.DataMember      =   "TAB_TP_CONTA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker txt_DT 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   609
      _Version        =   393216
      Format          =   65077249
      CurrentDate     =   38432
   End
   Begin MSDataListLib.DataCombo txt_NFicha 
      Bindings        =   "frm_Cad_Desc_Calc.frx":2FE2
      Height          =   360
      Left            =   4200
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "m_Nficha"
      BoundColumn     =   "M_F_COD"
      Text            =   ""
      Object.DataMember      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR"
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
      Left            =   3960
      TabIndex        =   13
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA"
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
      TabIndex        =   12
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTA"
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
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "FUNCIONÁRIO"
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
      TabIndex        =   9
      Top             =   1440
      Width           =   1455
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
      Left            =   360
      TabIndex        =   8
      Top             =   2880
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
      Left            =   4440
      TabIndex        =   6
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   3615
      Left            =   120
      Top             =   600
      Width           =   5175
   End
End
Attribute VB_Name = "frm_Cad_Desc_Calc"
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
Private Sub Cancelar()
On Error GoTo err1
    TXT_DESC = ""
    TXT_OP = ""
    txt_date = Date
    TXT_FUNC = ""
    txt_valor = 0
    TXT_CONTA = ""
    
    txt_DT.SetFocus
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Fechar()
On Error GoTo err1
    Unload Me
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Salvar()
On Error GoTo err1
    If (CDbl(Format(txt_DT, "mm")) = CDbl(Format(Date, "mm")) Or CDbl(Format(txt_DT, "mm")) = CDbl(Format(Date, "mm")) - 1) And (Not txt_valor = 0 Or txt_valor <> "") And TXT_OP <> "" And TXT_FUNC <> "" Then
        
        de.cmdIncluirDescCalc txt_DT, txt_NFicha.Text, TXT_CONTA.BoundText, TXT_OP, txt_valor, TXT_DESC,
        
        MsgBox "Registro salvo com sucesso!", vbInformation
        Cancelar
    ElseIf Not (CDbl(Format(txt_DT, "mm")) = CDbl(Format(Date, "mm")) Or CDbl(Format(txt_DT, "mm")) = CDbl(Format(Date, "mm")) - 1) Then
        MsgBox "Só é permitido data do mês passado ou do mês atual!", vbExclamation
    Else
        MsgBox "Preencha os Campos!", vbCritical
    End If
sair:
    Exit Sub
err1:
    If Err.Number = -2147467259 Then
        MsgBox "Este item já foi incluído na ficha!", vbExclamation
    Else
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    End If
    Resume sair
    
End Sub





Private Sub Form_Activate()
   Cancelar
End Sub

Private Sub Form_Load()
    Set adoFunc.Recordset = de.cnc.Execute("SELECT TAB_FICHA_MENS.M_F_COD, TAB_FUNCIONARIO.F_NOME, TAB_FICHA_MENS.M_NFICHA FROM TAB_FUNCIONARIO, TAB_FICHA_MENS WHERE TAB_FUNCIONARIO.F_Codigo = TAB_FICHA_MENS.M_F_COD AND (TAB_FICHA_MENS.M_BLOQ = 0)").Clone
    txt_DT = Date
    de.rsTAB_TP_CONTA.Requery
    TXT_CONTA.ReFill
    TXT_OP_CONTA.ReFill
End Sub





Private Sub TXT_CONTA_Change()
    TXT_OP_CONTA.BoundText = TXT_CONTA.BoundText
    TXT_OP = TXT_OP_CONTA.Text
End Sub

Private Sub TXT_FUNC_Change()
    txt_NFicha.BoundText = TXT_FUNC.BoundText
End Sub


'--------- Ao Pressionar uma Tecla -----------
Private Sub TXT_Conta_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub txt_DESC_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_OP_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_FUNC_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_dt_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_valor_KeyUp(KeyCode As Integer, Shift As Integer)
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

