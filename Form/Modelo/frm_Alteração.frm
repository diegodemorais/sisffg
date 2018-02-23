VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_Alteração 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VISUALIZAÇÃO COMPLETA"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   Icon            =   "frm_Alteração.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      DataField       =   "tp_cod"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
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
      TabIndex        =   5
      Top             =   960
      Width           =   810
   End
   Begin VB.ComboBox TXT_OP 
      DataField       =   "tp_op"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
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
      ItemData        =   "frm_Alteração.frx":1CFA
      Left            =   4440
      List            =   "frm_Alteração.frx":1D04
      TabIndex        =   2
      Text            =   "+"
      Top             =   1590
      Width           =   615
   End
   Begin VB.TextBox TXT_DESC 
      DataField       =   "tp_desc"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
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
      TabIndex        =   1
      Top             =   1590
      Width           =   3855
   End
   Begin MSDataGridLib.DataGrid Grid 
      Bindings        =   "frm_Alteração.frx":1D0E
      Height          =   5775
      Left            =   5520
      TabIndex        =   0
      Top             =   600
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   10186
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   -2147483639
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Caption         =   "TIPO DE CONTAS"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "TP_COD"
         Caption         =   "CÓD"
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
         DataField       =   "TP_DESC"
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
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   450,142
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   2340,284
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADOREG 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6465
      Width           =   9015
      _ExtentX        =   15901
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
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alteração.frx":1D23
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alteração.frx":203D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alteração.frx":2357
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alteração.frx":2671
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alteração.frx":298B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alteração.frx":2CA5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fechar"
            Object.ToolTipText     =   "Fechar (Alt+F)"
            ImageIndex      =   1
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
            Object.ToolTipText     =   "Editar Alteração (Alt+E)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "salvar"
            Object.ToolTipText     =   "Salvar Alteração (Alt+S)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar Alteração (Alt+C)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "excluir"
            Object.ToolTipText     =   "Excluir registro (Alt+X)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "filtrar"
            Object.ToolTipText     =   "Filtrar (Alt+T)"
            ImageIndex      =   6
         EndProperty
      EndProperty
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
      TabIndex        =   6
      Top             =   720
      Width           =   975
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
      TabIndex        =   4
      Top             =   1350
      Width           =   495
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
      TabIndex        =   3
      Top             =   1350
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   2055
      Left            =   120
      Top             =   600
      Width           =   5175
   End
End
Attribute VB_Name = "frm_Alteração"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W_LD_FILTRO As Boolean


Private Sub Form_Load()
On Error GoTo err1
    
    If de.rsTAB_TP_CONTA.State = 1 Then de.rsTAB_TP_CONTA.Close
    de.TAB_TP_CONTA
    Set ADOREG.Recordset = de.rsTAB_TP_CONTA.Clone
    ADOREG.Refresh
    de.rsTAB_TP_CONTA.Close

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

'*** Caption no navegador ***
Private Sub ADOREG_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1
    ADOREG.Caption = "REGISTRO : " & ADOREG.Recordset.AbsolutePosition & " / " & ADOREG.Recordset.RecordCount & IIf(W_LD_FILTRO = True, " (FILTRADO)", "")
sair:
    Exit Sub
err1:
    If Not Err.Number = -2147217885 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


'** Barra de Ferramenta ***
Private Sub BarraF_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
        Case "fechar": Fechar
        Case "editar": Editar
        Case "salvar": Salvar
        Case "cancelar": Cancelar
        Case "excluir": Excluir
        Case "filtrar": FILTRAR
    End Select
End Sub


'*** Rotinas ***
Private Sub Cancelar()
On Error GoTo err1
    
    POS = ADOREG.Recordset.AbsolutePosition - 1
    ADOREG.Recordset.CancelBatch adAffectCurrent
    ADOREG.Refresh
    ADOREG.Recordset.Move POS

    Editar
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub Editar()
On Error GoTo err1
    
    BarraF.Buttons("salvar").Enabled = Not BarraF.Buttons("salvar").Enabled
    BarraF.Buttons("cancelar").Enabled = Not BarraF.Buttons("cancelar").Enabled
    BarraF.Buttons("editar").Enabled = Not BarraF.Buttons("editar").Enabled
    
    Grid.Enabled = Not Grid.Enabled
    
    TXT_DESC.Enabled = Not TXT_DESC.Enabled
    TXT_OP.Enabled = Not TXT_OP.Enabled

    If BarraF.Buttons("salvar").Enabled = False Then
        Grid.SetFocus
    Else
        TXT_DESC.SetFocus
    End If

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Excluir()
On Error GoTo err1
        
    If vbYes = MsgBox("DESEJA REALMENTE EXCLUIR O TIPO DE CONTA (" & TXT_DESC & ")?", vbQuestion + vbYesNo) Then
        ADOREG.Recordset.Delete
        ADOREG.Recordset.UpdateBatch
    End If
    
sair:
    Exit Sub
err1:
    If Not Err.Number = -2147467259 Then
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    Else
        MsgBox "NÃO É POSSÍVEL EXCLUIR ESTE TIPO DE CONTA, DEVIDO A VENDAS RELACIONADAS A ELE!", vbCritical
        ADOREG.Refresh
    End If
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


Private Sub FILTRAR()
Dim w_resp As String
Dim W_CAMPO As String
Dim W_FILTRO As String

On Error GoTo err1
    
    w_resp = InputBox("FILTRAR PELO QUÊ ? ENTRE COM O NÚMERO DA OPÇÃO DESEJADA." & Chr(13) & Chr(13) & "1 - DESCRIÇÃO" & Chr(13) & "2 - OPERAÇÃO" & Chr(13) & "3 - REMOVER FILTRO *", , "1")
    
    
    If Not w_resp = "" And IsNumeric(w_resp) And w_resp >= 1 And w_resp <= 5 Then
        Select Case w_resp
        'DESC
        Case 1:
            w_resp = "DESCRIÇÃO"
            W_CAMPO = "TP_DESC"
        'OP
        Case 2:
            w_resp = "OPERAÇÃO"
            W_CAMPO = "TP_OP"
            
        '*** REMOVE O FILTRO ****
        Case 3:
            If Not ADOREG.Recordset.Filter = 0 Then
                W_LD_FILTRO = False
                ADOREG.Recordset.Filter = 0
                ADOREG.Refresh
            End If
        End Select
        
        If Not w_resp = "3" Then
            W_FILTRO = InputBox("ENTRE COM O " & w_resp & " DESEJADO !")
            If Not W_FILTRO = "" Then
                W_FILTRO = W_CAMPO & " LIKE '%" & W_FILTRO & "%'"
                W_LD_FILTRO = True
                ADOREG.Recordset.Filter = W_FILTRO
            End If
        End If
    End If
    
sair:
    Exit Sub
err1:
    If Err.Number <> 13 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
        W_LD_FILTRO = False
        Resume sair

End Sub

Private Sub Salvar()
On Error GoTo err1
    
    ADOREG.Recordset.UpdateBatch adAffectCurrent
    
    Editar
    
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
    Case 69: ' "E"
            Editar
    Case 83: ' "S"
            Salvar
    Case 67: ' "C"
            Cancelar
    Case 88: ' "X"
            Excluir
    Case 84: ' "T"
            FILTRAR
    End Select
End If
End Sub



