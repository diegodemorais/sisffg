VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "msCOMCTL.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.6#0"; "ACTIVETEXT.OCX"
Begin VB.Form frm_Log 
   AutoRedraw      =   -1  'True
   Caption         =   "Auditoria"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   13905
   Icon            =   "frm_Log.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   13905
   StartUpPosition =   1  'CenterOwner
   Begin rdActiveText.ActiveText txtDtFim 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   8
      TextMask        =   9
      RawText         =   9
      DateFormat      =   1
      Mask            =   "##/##/##"
      FontBold        =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
   End
   Begin rdActiveText.ActiveText txtDtIni 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   8
      TextMask        =   9
      RawText         =   9
      DateFormat      =   1
      Mask            =   "##/##/##"
      FontBold        =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
   End
   Begin VB.TextBox txtDescricao 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6840
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Width           =   5535
   End
   Begin VB.ComboBox cbUsuario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frm_Log.frx":030A
      Left            =   3720
      List            =   "frm_Log.frx":0320
      TabIndex        =   2
      Text            =   "TODOS"
      Top             =   1320
      Width           =   1380
   End
   Begin VB.CommandButton cmdPesq 
      Caption         =   "&Buscar"
      Height          =   855
      Left            =   12600
      Picture         =   "frm_Log.frx":034C
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   975
   End
   Begin VB.ComboBox cbAcao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frm_Log.frx":31C6
      Left            =   5280
      List            =   "frm_Log.frx":31D6
      TabIndex        =   3
      Text            =   "TODAS"
      Top             =   1320
      Width           =   1380
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Log.frx":31FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Log.frx":3515
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Log.frx":382F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Log.frx":3B49
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Log.frx":3E63
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Log.frx":417D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Log.frx":4497
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   1482
      ButtonWidth     =   1376
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Fechar"
            Key             =   "fechar"
            Object.ToolTipText     =   "Fechar (Alt+F)"
            ImageIndex      =   1
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Detalhar"
            Key             =   "detalhar"
            Description     =   "Detalhar registro"
            Object.ToolTipText     =   "Detalhar registro"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoConta 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   8640
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   2
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
      Caption         =   "REGISTRO : 0/0"
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
   Begin MSDataGridLib.DataGrid grid_Conta 
      Bindings        =   "frm_Log.frx":48E9
      Height          =   6450
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2160
      Width           =   13770
      _ExtentX        =   24289
      _ExtentY        =   11377
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "log_codigo"
         Caption         =   "COD"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "log_data"
         Caption         =   "DATA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "log_hora"
         Caption         =   "HORA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   4
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "log_usuario"
         Caption         =   "USUÁRIO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "log_acao"
         Caption         =   "AÇÃO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "log_tabela"
         Caption         =   "TABELA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "log_descricao"
         Caption         =   "DESCRIÇÃO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTENDO"
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
      Left            =   6840
      TabIndex        =   13
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Entre"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "e"
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "USUÁRIO"
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
      Left            =   3720
      TabIndex        =   10
      Top             =   1080
      Width           =   975
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
      TabIndex        =   9
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "AÇÃO"
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
      Left            =   5280
      TabIndex        =   8
      Top             =   1080
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   120
      Top             =   840
      Width           =   13770
   End
End
Attribute VB_Name = "frm_Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'** Barra de Ferramenta ***
Private Sub BarraF_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
        Case "fechar": Fechar
        Case "detalhar": MostrarDetalhes
    End Select
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

Private Sub cbAcao_Click()
    cmdPesq_Click
End Sub

Private Sub cbUsuario_Click()
    cmdPesq_Click
End Sub

Sub cmdPesq_Click()
Dim w_acaoLog, w_usuarioLog As String

    
    If cbAcao = "TODAS" Then w_acaoLog = "%" Else w_acaoLog = cbAcao
    If cbUsuario = "TODOS" Then w_usuarioLog = "%" Else w_usuarioLog = cbUsuario

    If de.rscmdBase.State = 1 Then de.rscmdBase.Close
    If adoConta.Recordset.State = 1 Then adoConta.Recordset.Close
    
    de.rscmdBase.Open "SELECT log_codigo, log_data, log_hora,  log_usuario , log_acao, log_tabela, log_descricao FROM TAB_LOG WHERE (log_data >= #" & Format(CVDate(txtDtIni), "mm/dd/YYYY") & "#) AND (log_data <= #" & Format(CVDate(txtDtFim), "mm/dd/YYYY") & "#) AND (log_usuario like '%" & w_usuarioLog & "%') AND (log_acao like '%" & w_acaoLog & "%') AND (log_descricao LIKE '%" & Trim(txtDescricao) & "%' OR log_tabela LIKE '%" & Trim(txtDescricao) & "%') ", , adOpenStatic, adLockOptimistic
    Set adoConta.Recordset = de.rscmdBase.Clone
    de.rscmdBase.Close
    'adoConta.Recordset.Close
    

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub Form_Load()
On erro GoTo err1
Dim w_acaoLog, w_usuarioLog As String

    txtDtIni = Format(CDate(Now()), "dd/mm/yy")
    txtDtFim = Format(CDate(Now()), "dd/mm/yy")
    
    If cbAcao = "TODAS" Then w_acaoLog = "%" Else w_acaoLog = cbAcao
    If cbUsuario = "TODOS" Then w_usuarioLog = "%" Else w_usuarioLog = cbUsuario

    If de.rscmdBase.State = 1 Then de.rscmdBase.Close
    
    de.rscmdBase.Open "SELECT log_codigo, log_data, log_hora,  log_usuario , log_acao, log_tabela, log_descricao FROM TAB_LOG WHERE (log_data >= #" & Format(CVDate(txtDtIni), "mm/dd/YYYY") & "#) AND (log_data <= #" & Format(CVDate(txtDtFim), "mm/dd/YYYY") & "#) AND (log_usuario like '%" & w_usuarioLog & "%') AND (log_acao like '%" & w_acaoLog & "%') AND (log_descricao LIKE '%" & Trim(txtDescricao) & "%' OR log_tabela LIKE '%" & Trim(txtDescricao) & "%') ", , adOpenStatic, adLockOptimistic
    Set adoConta.Recordset = de.rscmdBase.Clone
    de.rscmdBase.Close

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Sub MostrarDetalhes()
Dim w_DescricaoLog() As String
    
    'w_DescricaoLog(1) = ""
    w_DescricaoLog = Split(adoConta.Recordset.Fields("log_descricao"), ">>> ")
    
    frm_LogDet.txtCod = adoConta.Recordset.Fields("log_codigo")
    frm_LogDet.txtData = adoConta.Recordset.Fields("log_data")
    frm_LogDet.txtHora = adoConta.Recordset.Fields("log_hora")
    frm_LogDet.txtUsuario = adoConta.Recordset.Fields("log_usuario")
    frm_LogDet.txtAcao = adoConta.Recordset.Fields("log_acao")
    frm_LogDet.txtTabela = adoConta.Recordset.Fields("log_tabela")
    frm_LogDet.txtDescricaoNEW = w_DescricaoLog(0)
    
    If adoConta.Recordset.Fields("log_acao") = "EDITAR" And UBound(w_DescricaoLog) >= 1 Then
        frm_LogDet.lblDescricaoOLD.Visible = True
        frm_LogDet.txtDescricaoOLD.Visible = True
        frm_LogDet.txtDescricaoOLD = w_DescricaoLog(1)
    Else
        frm_LogDet.lblDescricaoOLD.Visible = False
        frm_LogDet.txtDescricaoOLD.Visible = False
    End If
    
    frm_LogDet.Show 1
    
    'Text1.Text = "COD: " & adoConta.Recordset.Fields("log_codigo") & " DATA: " & adoConta.Recordset.Fields("log_data") & " HORA: " & adoConta.Recordset.Fields("log_hora") & " USUÁRIO: " & adoConta.Recordset.Fields("log_usuario") & " AÇÃO: " & adoConta.Recordset.Fields("log_acao") & " TABELA: " & adoConta.Recordset.Fields("log_tabela") & " DESCRIÇÃO: " & adoConta.Recordset.Fields("log_descricao")
    'Text1.Visible = True
    'Text1.SetFocus
End Sub


Private Sub grid_Conta_DblClick()
    MostrarDetalhes
End Sub

Private Sub grid_Conta_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            MostrarDetalhes
    End Select
End Sub

'--------- Ao Pressionar uma Tecla -----------

Private Sub grid_Conta_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub GRID_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub


Private Sub TXT_CONTA_Change()
    TXT_CONTA_cod = TXT_CONTA.BoundText
End Sub

Private Sub TXT_CONTA_COD_Change()
    'TXT_CONTA.BoundText = TXT_CONTA_cod
End Sub


'*** KEYASCII ***
Private Sub TXT_ANO_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then
        Sendkeys "{tab}"
        cmdPesq_Click
      End If
End Sub
Private Sub TXT_CONTA_COD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Sendkeys "{tab}"
    End If
End Sub

Private Sub TXT_CONTA_cod_LostFocus()
    TXT_CONTA.BoundText = TXT_CONTA_cod
    cmdPesq_Click
End Sub

Private Sub TXT_CONTA_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 And ckTodas.value = 0 Then
        TXT_LOGO.SetFocus
        cmdPesq_Click
      ElseIf KeyCode = 13 Then
        ckTodas.SetFocus
        cmdPesq_Click
      End If
End Sub

Private Sub TXT_LOGO_Change()
   If TXT_LOGO <> "" Then ck_Nome.value = 1
End Sub

Private Sub TXT_LOGO_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then
        Sendkeys "{tab}"
        cmdPesq_Click
      End If
End Sub
Private Sub TXT_MES_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then
        Sendkeys "{tab}"
        cmdPesq_Click
      End If

End Sub



Private Sub TXT_total_GotFocus()
    grid_Conta.SetFocus
End Sub





' -------  Teclas de Atalhos --------

Sub Keys(KeyCode As Integer, Shift As Integer)
'*** Shift (4 = Alt) ***
If Shift = 4 Then
    Select Case KeyCode
    Case 70: ' "F"
            Fechar
    End Select
End If
End Sub


Private Sub txtDescricao_Change()
    cmdPesq_Click
End Sub

Private Sub txtDtFim_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then cmdPesq_Click
End Sub

Private Sub txtDtIni_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then
        txtDtFim.SetFocus
      End If
End Sub
