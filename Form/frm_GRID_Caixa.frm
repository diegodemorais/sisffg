VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_GRID_Caixa 
   Caption         =   "Comissão Caixas"
   ClientHeight    =   10650
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   ScaleHeight     =   10650
   ScaleWidth      =   12465
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBrt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$"" #.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
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
      Left            =   8925
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   9885
      Width           =   1500
   End
   Begin VB.TextBox txtLiq 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$"" #.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
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
      Left            =   10485
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   9885
      Width           =   1500
   End
   Begin VB.TextBox txtMinimo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$"" #.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
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
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   9885
      Width           =   1500
   End
   Begin VB.TextBox txtFixo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$"" #.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
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
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   9885
      Width           =   1500
   End
   Begin Skin_Button.ctr_Button btnRptGRIDcaixa 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Relatório dos @"
      Top             =   0
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   503
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   12632319
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_GRID_Caixa.frx":0000
      PICN            =   "frm_GRID_Caixa.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataGridLib.DataGrid Grid 
      Bindings        =   "frm_GRID_Caixa.frx":12FE
      Height          =   9495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   16748
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   18
      TabAction       =   1
      WrapCellPointer =   -1  'True
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "SIGLA"
         Caption         =   "Sigla"
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
         DataField       =   "B"
         Caption         =   "B"
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
      BeginProperty Column02 
         DataField       =   "F_TIPO"
         Caption         =   "Cx"
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
      BeginProperty Column03 
         DataField       =   "NOME"
         Caption         =   "Nome"
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
         DataField       =   "FIXO"
         Caption         =   "Fixo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "MINIMO"
         Caption         =   "Mínimo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "COMISSAO"
         Caption         =   "< 5%"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   5
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "COMISSAO2"
         Caption         =   "< 10%"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   5
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "COMISSAO3"
         Caption         =   "> 10%"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   5
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "PERC_DEZ"
         Caption         =   "+ Dezembro"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   5
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "TIPO"
         Caption         =   "TIPO"
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
      BeginProperty Column11 
         DataField       =   "B"
         Caption         =   "B"
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
      BeginProperty Column12 
         DataField       =   "M_NFICHA"
         Caption         =   "M_NFICHA"
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
      BeginProperty Column13 
         DataField       =   "F_VPISO"
         Caption         =   "Piso Brt"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "F_VPISO_R"
         Caption         =   "Piso Líq"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            Alignment       =   2
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column13 
            Alignment       =   1
         EndProperty
         BeginProperty Column14 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoReg 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   10320
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   582
      ConnectMode     =   3
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
      Caption         =   "Registro(s):"
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
   Begin VB.Shape Shape3 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   465
      Left            =   8880
      Top             =   9840
      Width           =   3180
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Piso Brt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   9240
      TabIndex        =   9
      Top             =   9525
      Width           =   825
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Piso Líq"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   10800
      TabIndex        =   8
      Top             =   9525
      Width           =   840
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mínimo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   7560
      TabIndex        =   4
      Top             =   9525
      Width           =   735
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fixo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   6240
      TabIndex        =   3
      Top             =   9525
      Width           =   450
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   465
      Left            =   5595
      Top             =   9840
      Width           =   3180
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   345
      Left            =   5595
      Top             =   9480
      Width           =   3180
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   345
      Left            =   8880
      Top             =   9480
      Width           =   3180
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "Opções"
      Index           =   0
      Begin VB.Menu mnuNotas 
         Caption         =   "Notas"
         Index           =   1
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frm_GRID_Caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wMesAtual, wAnoAtual, wMaxFicha As String

Private Sub btnRptGRIDcaixa_Click()
On Error GoTo err1
    If de.rscmdSqlCaixaComissao2_Grouping.State = 1 Then de.rscmdSqlCaixaComissao2_Grouping.Close
    de.cmdSqlCaixaComissao2_Grouping wMesAtual, wAnoAtual
    
    rptGRID_Caixa.Show 1
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    wMaxFicha = de.cnc.Execute("SELECT Max(M_NFICHA) FROM TAB_FICHA_MENS").Fields(0)
    wMesAtual = de.cnc.Execute("SELECT Max(M_MES) FROM TAB_FICHA_MENS WHERE M_NFICHA = " & wMaxFicha).Fields(0)
    wAnoAtual = de.cnc.Execute("SELECT Max(M_ANO) FROM TAB_FICHA_MENS WHERE M_NFICHA = " & wMaxFicha).Fields(0)
    
    
    de.cnc.CursorLocation = adUseServer
    If de.rscmdSqlCaixaComissao.State = 1 Then de.rscmdSqlCaixaComissao.Close
    de.cmdSqlCaixaComissao wMesAtual, wAnoAtual
    
    Set adoReg.Recordset = de.rscmdSqlCaixaComissao.Clone
    
    Call CalcGrid
End Sub

Sub CalcGrid()
    Dim adoGrid As ADODB.Recordset
    Dim wFixo, wMinimo, wBrt, wLiq

    wFixo = 0
    wMinimo = 0

    Set adoGrid = adoReg.Recordset.Clone
    If adoGrid.RecordCount <> 0 Then adoGrid.MoveFirst
    Do While Not adoGrid.EOF
        If Not IsNull(adoGrid.Fields("FIXO")) Then wFixo = wFixo + CDbl(adoGrid.Fields("FIXO"))
        If Not IsNull(adoGrid.Fields("MINIMO")) Then wMinimo = wMinimo + CDbl(adoGrid.Fields("MINIMO"))
        If Not IsNull(adoGrid.Fields("F_VPISO")) Then wBrt = wBrt + CDbl(adoGrid.Fields("F_VPISO"))
        If Not IsNull(adoGrid.Fields("F_VPISO_R")) Then wLiq = wLiq + CDbl(adoGrid.Fields("F_VPISO_R"))

        adoGrid.MoveNext
    Loop
    
    txtFixo = Format(wFixo, "R$ #,##0.00")
    txtMinimo = Format(wMinimo, "R$ #,##0.00")
    txtBrt = Format(wBrt, "R$ #,##0.00")
    txtLiq = Format(wLiq, "R$ #,##0.00")
    
End Sub


Private Sub Grid_AfterUpdate()
    Call CalcGrid
End Sub



Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Sendkeys "{tab}"
    End If
End Sub

Private Sub mnuNotas_Click(Index As Integer)
    Call MostrarDetalhes
End Sub

Sub MostrarDetalhes()
Dim wtxtNotas
    wtxtNotas = de.cnc.Execute("SELECT F_NOTAS FROM TAB_FUNCIONARIO WHERE F_Codigo = " & adoReg.Recordset.Fields("F_Codigo")).Fields(0)
    If IsNull(wtxtNotas) Then wtxtNotas = ""
    
    frm_GRID_Caixa_Det.txtSigla = adoReg.Recordset.Fields("SIGLA")
    frm_GRID_Caixa_Det.txtB = adoReg.Recordset.Fields("B")
    frm_GRID_Caixa_Det.txtTipo = adoReg.Recordset.Fields("TIPO")
    frm_GRID_Caixa_Det.txtNome = adoReg.Recordset.Fields("NOME")
    frm_GRID_Caixa_Det.txtFunc = adoReg.Recordset.Fields("F_Codigo")
    frm_GRID_Caixa_Det.txtFicha = adoReg.Recordset.Fields("M_NFICHA")
    frm_GRID_Caixa_Det.txtNotas = wtxtNotas
    frm_GRID_Caixa_Det.txtNotasOld = wtxtNotas
    frm_GRID_Caixa_Det.Show 1

End Sub
