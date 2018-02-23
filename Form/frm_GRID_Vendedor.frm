VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_GRID_Vendedor 
   Caption         =   "Códigos Vendedores"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ckTodas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CommandButton cmdPesq 
      Caption         =   "&Buscar"
      Height          =   735
      Left            =   4200
      Picture         =   "frm_GRID_Vendedor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid Grid 
      Bindings        =   "frm_GRID_Vendedor.frx":2E7A
      Height          =   7680
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   13547
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "SIGLA"
         Caption         =   "SIGLA"
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
         DataField       =   "NOME"
         Caption         =   "NOME"
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
         DataField       =   "COD_MILLENNIUM"
         Caption         =   "CÓDIGO MILLENNIUM"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "M_ANO"
         Caption         =   "M_ANO"
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
         DataField       =   "M_MES"
         Caption         =   "M_MES"
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
         DataField       =   "F_Codigo"
         Caption         =   "F_Codigo"
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
            Alignment       =   2
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3644,788
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1904,882
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   0
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoReg 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   8910
      Width           =   7335
      _ExtentX        =   12938
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
   Begin MSDataListLib.DataCombo TXT_LOGO 
      Bindings        =   "frm_GRID_Vendedor.frx":2E8F
      DataField       =   "F_COD_L"
      DataSource      =   "ADOREG2"
      Height          =   360
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "COD_LOJ"
      BoundColumn     =   "COD_LOJ"
      Text            =   "%"
      Object.DataMember      =   "TAB_L"
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
   Begin MSDataListLib.DataCombo TXT_LOGO2 
      Bindings        =   "frm_GRID_Vendedor.frx":2EA0
      DataField       =   "F_COD_L"
      DataSource      =   "ADOREG2"
      Height          =   360
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "NUM"
      BoundColumn     =   "COD_LOJ"
      Text            =   "%"
      Object.DataMember      =   "TAB_L_NUM"
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "(B)"
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
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   0
      Top             =   120
      Width           =   7290
   End
End
Attribute VB_Name = "frm_GRID_Vendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ano, mes As String

Sub ckTodas_Click()
    If ckTodas.value = 1 Then
        TXT_LOGO = "%"
        TXT_LOGO.Enabled = False
        TXT_LOGO2 = "%"
        TXT_LOGO2.Enabled = False
    Else
        TXT_LOGO = ""
        TXT_LOGO.Enabled = True
        TXT_LOGO2 = ""
        TXT_LOGO2.Enabled = True
        On Error Resume Next
        TXT_LOGO2.SetFocus
        SendKeys "{f4}"
    End If
End Sub

Private Sub cmdPesq_Click()

    If de.rscmdSqlGRIDVendedor.State = 1 Then de.rscmdSqlGRIDVendedor.Close
    de.cmdSqlGRIDVendedor ano, mes
    
    Set adoReg.Recordset = de.rscmdSqlGRIDVendedor.Clone

    If TXT_LOGO = "" Or ckTodas.value = 1 Then
        ckTodas.value = 1
        adoReg.Recordset.Filter = 0
    Else
        adoReg.Recordset.Filter = "SIGLA = '" & TXT_LOGO.Text & "'"
    End If
End Sub

Private Sub Form_Load()
      
    ano = InputBox("Digite o ano: ", , Format(Date, "YYYY"))
    mes = InputBox("Digite o mês: ", , Format(Date, "MM"))
    
    If de.rscmdSqlGRIDVendedor.State = 1 Then de.rscmdSqlGRIDVendedor.Close
    de.cmdSqlGRIDVendedor ano, mes
    
    Set adoReg.Recordset = de.rscmdSqlGRIDVendedor.Clone
    
    ckTodas_Click
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub TXT_LOGO_Change()
    TXT_LOGO2.BoundText = TXT_LOGO.BoundText
End Sub



Private Sub TXT_LOGO2_Change()
    TXT_LOGO.BoundText = TXT_LOGO2.BoundText
End Sub


