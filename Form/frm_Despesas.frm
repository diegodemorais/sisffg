VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_Despesas 
   Caption         =   "Despesas Extras do Caixa - Exportação"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16425
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   16425
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid grid_Off 
      Bindings        =   "frm_Despesas.frx":0000
      Height          =   5010
      Left            =   1800
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   8837
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   12640511
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
      Caption         =   "E X C L U Í D O S"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "LANCAMENTO_FILIAL_FILIAL_NOME"
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
      BeginProperty Column01 
         DataField       =   "F_1907459465"
         Caption         =   "DATA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "LANCAMENTO_LANCAMENTO_OBS"
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
      BeginProperty Column03 
         DataField       =   "F_2282248934"
         Caption         =   "F_2282248934"
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
         DataField       =   "F_3518719985"
         Caption         =   "F_3518719985"
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
         DataField       =   "SUM_LANCAMENTO_VALOR_INICIAL_"
         Caption         =   "SUM_LANCAMENTO_VALOR_INICIAL_"
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
         DataField       =   "F_2714866558"
         Caption         =   "VALOR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            Alignment       =   2
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoOff 
      Height          =   330
      Left            =   1800
      Top             =   6720
      Visible         =   0   'False
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   2
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
   Begin VB.TextBox txtLog 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3285
      Left            =   12360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Top             =   5640
      Width           =   3930
   End
   Begin VB.Frame frmExportar 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Exportar para Sistema de Contas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   12360
      TabIndex        =   14
      Top             =   2280
      Width           =   3975
      Begin VB.TextBox txtB 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1680
         Width           =   1380
      End
      Begin VB.TextBox txtValor 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Height          =   405
         Left            =   1440
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "R$ 0,00"
         Top             =   1080
         Width           =   1380
      End
      Begin Skin_Button.ctr_Button cmdExportar 
         Height          =   525
         Left            =   2400
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "Exportar"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_Despesas.frx":0015
         PICN            =   "frm_Despesas.frx":0031
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker txtData 
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   480
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   216399873
         CurrentDate     =   38282
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "B:"
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
         Left            =   960
         TabIndex        =   21
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
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
         Left            =   600
         TabIndex        =   20
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
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
         Left            =   720
         TabIndex        =   17
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdDel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   15960
      Picture         =   "frm_Despesas.frx":0483
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Default         =   -1  'True
      Height          =   375
      Left            =   10560
      Picture         =   "frm_Despesas.frx":078D
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   480
      Width           =   375
   End
   Begin VB.ListBox lstCriterios 
      BackColor       =   &H00C0C0FF&
      DataSource      =   "adoCriterios"
      Height          =   1620
      ItemData        =   "frm_Despesas.frx":0A97
      Left            =   12360
      List            =   "frm_Despesas.frx":0A99
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox txtCriterio 
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   480
      Width           =   6975
   End
   Begin VB.CommandButton cmdPesq 
      Caption         =   "&Buscar"
      Height          =   735
      Left            =   11280
      Picture         =   "frm_Despesas.frx":0A9B
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox TXT_TOTAL 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      Height          =   330
      Left            =   14115
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   9075
      Width           =   1620
   End
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
      Left            =   2040
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc adoDespesas 
      Height          =   330
      Left            =   240
      Top             =   9000
      Visible         =   0   'False
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   2
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
   Begin MSDataGridLib.DataGrid grid_Despesas 
      Bindings        =   "frm_Despesas.frx":3915
      Height          =   8490
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   14975
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
         DataField       =   "LANCAMENTO_FILIAL_FILIAL_NOME"
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
      BeginProperty Column01 
         DataField       =   "F_1907459465"
         Caption         =   "DATA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "LANCAMENTO_LANCAMENTO_OBS"
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
      BeginProperty Column03 
         DataField       =   "F_2282248934"
         Caption         =   "F_2282248934"
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
         DataField       =   "F_3518719985"
         Caption         =   "F_3518719985"
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
         DataField       =   "SUM_LANCAMENTO_VALOR_INICIAL_"
         Caption         =   "SUM_LANCAMENTO_VALOR_INICIAL_"
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
         DataField       =   "F_2714866558"
         Caption         =   "VALOR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            Alignment       =   2
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo TXT_LOGO 
      Bindings        =   "frm_Despesas.frx":392F
      DataField       =   "F_COD_L"
      DataSource      =   "ADOREG"
      Height          =   360
      Left            =   150
      TabIndex        =   3
      Top             =   480
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
      Bindings        =   "frm_Despesas.frx":3940
      DataField       =   "F_COD_L"
      DataSource      =   "ADOREG"
      Height          =   360
      Left            =   1080
      TabIndex        =   4
      Top             =   480
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
   Begin VB.Label lbCount 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   15960
      TabIndex        =   24
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Histórico de exportação:"
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
      Left            =   12360
      TabIndex        =   23
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Critérios:"
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
      Left            =   12360
      TabIndex        =   13
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Adicionar Critério:"
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
      Left            =   3480
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL :"
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
      Index           =   12
      Left            =   12990
      TabIndex        =   6
      Top             =   9165
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "(B):"
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
      Left            =   150
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   12750
      Top             =   9000
      Width           =   3135
   End
End
Attribute VB_Name = "frm_Despesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wTotal As Double
Dim wSQLini, wSQLfim, wSQL, wSQLcritINI, wSQLcrit, wSQLCOUNT, wSQLcritCOUNT As String
Dim wCount As Integer

Private Sub ckTodas_Click()
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
        Sendkeys "{f4}"
    End If
End Sub

Private Sub cmdFiltrar_Click()
    lstCriterios.AddItem (txtCriterio.TabIndex)
End Sub

Private Sub cmdAdd_Click()
    If Trim(txtCriterio.text) <> "" Then
        lstCriterios.AddItem txtCriterio
        txtCriterio.text = ""
        Call list2txt(lstCriterios, "DespesasCriterios.txt")
        cmdPesq_Click
        'adoCriterio.Recordset.UpdateBatch adAffectCurrent
    End If
End Sub

Private Sub cmdDel_Click()
    If lstCriterios.ListIndex <> -1 Then
        lstCriterios.RemoveItem lstCriterios.ListIndex
        'adoCriterio.Recordset.UpdateBatch adAffectCurrent
        Call list2txt(lstCriterios, "DespesasCriterios.txt")
        cmdPesq_Click
    End If
End Sub

Private Sub cmdExportar_Click()
Dim lastNumProcesso
Dim wNumContas
    wNunContas = 99

   If MsgBox("Deseja exportar as Despesas Extras do Caixa na data de " & txtData & " para a " & txtB & "? (" & txtValor & ")", vbYesNo, "GERAR DESPESAS EXTRAS DO CAIXA") = vbYes Then
        'Despesas Extras do Caixa
        If de.cncContas.State = 0 Then de.cncContas.Open
        wNumContas = de.cncContas.Execute("SELECT tblloja_loja_contas FROM tblloja where tblloja_loja_fichas = '" & Format(txtB, "000") & "'").Fields(0)
        de.cmdAddProcesso "1595", CDate(Now()), "**Incluído automaticamente**", "R", CDbl(txtValor), wNumContas, CDate(Now())
        lastNumProcesso = de.cncContas.Execute("SELECT MAX(tblpro_num) FROM tblpropg").Fields(0)
        de.cmdAddProcessoItem lastNumProcesso, txtData, CDbl(txtValor), txtData
        txtLog = txtLog & "B: " & txtB & " Data: " & txtData & " Valor: " & txtValor & " Processo: " & lastNumProcesso & vbNewLine
        'FIM Despesas Extras do Caixa
   End If
End Sub

Sub cmdPesq_Click()
    loja = TXT_LOGO2

'    Select Case loja
'        Case "F.MORATO ZULIAN ME":
'            loja = "01"
'        Case "F. MORATO ZULIAN ME - FILIAL 1":
'            loja = "06"
'        Case "F. MORATO ZULIAN ME - FILIAL 2":
'            loja = "17"
'        Case "F. MORATO ZULIAN ME - FILIAL 3":
'            loja = "66"
'    End Select
    
    cmdAdd_Click
    
    If ckTodas Then
        wSQLcritINI = "WHERE LANCAMENTO_FILIAL_FILIAL_NOME <> '98'"
    Else
        wSQLcritINI = "WHERE LANCAMENTO_FILIAL_FILIAL_NOME = " & loja & " AND LANCAMENTO_FILIAL_FILIAL_NOME <> '98'"
        'adoDespesas.Recordset.Filter = "LANCAMENTO_FILIAL_FILIAL_NOME = " & TXT_LOGO2
    End If
    
    wSQLcrit = ""
    Dim I As Integer
    With lstCriterios
        For I = 0 To .ListCount - 1
        wSQLcrit = wSQLcrit & " AND UPPER(lancamento_lancamento_OBS) NOT LIKE '" & UCase(.list(I)) & "%'"
        'adoDespesas.Recordset.Filter = adoDespesas.Recordset.Filter & " AND (LANCAMENTO_LANCAMENTO_OBS < '" & .List(i) & "' AND LANCAMENTO_LANCAMENTO_OBS > '" & .List(i) & "')"
    
    Next
    End With

    'de.cncMwts.Open
       
    wSQL = wSQLini & wSQLcritINI & wSQLcrit & wSQLfim
    If de.cncMwts.State = 1 Then
        de.cncMwts.Close
        de.cncMwts.Open
    End If
    Set adoDespesas.Recordset = de.cncMwts.Execute(wSQL).Clone
    'de.cncMwts.Close
   
    wCount = 0
    wSQLcritCOUNT = " AND ( UPPER(lancamento_lancamento_OBS) = 'ZYZYZYZY'" & wSQLcrit
    wSQLcritCOUNT = Replace(wSQLcritCOUNT, " NOT ", " ")
    wSQLcritCOUNT = Replace(wSQLcritCOUNT, "AND UPPER", "OR UPPER")
    wSQLCOUNT = wSQLini & wSQLcritINI & wSQLcritCOUNT & ")" & wSQLfim
    Set adoOff.Recordset = de.cncMwts.Execute(wSQLCOUNT).Clone
    wCount = adoOff.Recordset.RecordCount
   
    wTotal = 0
    If adoDespesas.Recordset.RecordCount <> 0 Then adoDespesas.Recordset.MoveFirst
    Do While Not adoDespesas.Recordset.EOF
        wTotal = wTotal + CDbl(adoDespesas.Recordset.Fields("F_2714866558"))
        adoDespesas.Recordset.MoveNext
    Loop
    
    TXT_TOTAL = Format(wTotal, "R$ 0.00")
    
    If adoDespesas.Recordset.RecordCount <> 0 Then adoDespesas.Recordset.MoveFirst
    If adoDespesas.Recordset.RecordCount <= 0 Then MsgBox "Não existe DESPESAS nestas características!", vbExclamation
    
    txtValor.text = TXT_TOTAL.text
    txtB.text = loja
    
    cmdExportar.Enabled = Not ckTodas.value
    
    lbCount.Caption = wCount
    
    
   'Varrendo todas as linhas do flexgrid
'   For I = 1 To grid_Despesas_L.Rows - 1
'        grid_Despesas_L.TextMatrix(I, 1) = Format(grid_Despesas_L.TextMatrix(I, 1), "DD/MM/YYYY")
'        grid_Despesas_L.TextMatrix(I, 6) = FormatCurrency(Round(grid_Despesas_L.TextMatrix(I, 6), 2))
'        If grid_Despesas_L.TextMatrix(I, 6) > 99 Then
'             For coluna = 0 To grid_Despesas_L.Cols - 1
'                 grid_Despesas_L.Col = coluna
'                 grid_Despesas_L.Row = I
'                 grid_Despesas_L.CellFontBold = True
'                 grid_Despesas_L.CellForeColor = vbRed
'             Next coluna
'        ElseIf grid_Despesas_L.TextMatrix(I, 6) > 49 Then
'             For coluna = 0 To grid_Despesas_L.Cols - 1
'                 grid_Despesas_L.Col = coluna
'                 grid_Despesas_L.Row = I
'                 grid_Despesas_L.CellFontBold = True
'                 grid_Despesas_L.CellForeColor = vbBlue
'             Next coluna
'        End If
'   Next I


    
    
End Sub

Private Sub Form_Activate()

'    grid_Despesas_L.ColWidth(0) = 420 'B
'    grid_Despesas_L.ColWidth(1) = 1200 'data
'    grid_Despesas_L.ColWidth(2) = 8500 'descrição
'    grid_Despesas_L.ColWidth(3) = 0 'cod filial
'    grid_Despesas_L.ColWidth(4) = 0 'cod lançamento
'    grid_Despesas_L.ColWidth(5) = 0 'valor negativo
'    grid_Despesas_L.ColWidth(6) = 1080 'valor

'    grid_Despesas_L.ColAlignment(0) = flexAlignCenterBottom 'B
'    grid_Despesas_L.ColAlignment(6) = flexAlignRightBottom 'valor


End Sub

Private Sub Form_Load()
    'If de.rscmdDespesas.State = 1 Then de.rscmdDespesas.Close
    'de.cmdDespesas
    'Set adoDespesas.Recordset = de.rscmdDespesas.Clone
    If de.cncMwts.State = 0 Then de.cncMwts.Open
    wSQLini = "select v.lancamento_filial_filial_nome  as lancamento_filial_filial_nome , v.F_1907459465 as F_1907459465 ," _
        & " v.lancamento_lancamento_OBS  as lancamento_lancamento_OBS , v.F_2282248934  as F_2282248934 , v.F_3518719985 " _
        & " as F_3518719985 , coalesce ( sum ( v.sum_lancamento_VALOR_INICIAL_ ) , 0 )  as sum_lancamento_VALOR_INICIAL_ ," _
        & " coalesce ( sum ( v.F_2714866558 ) , 0 )  as F_2714866558 from ( select DISTINCT F_2463272603.COD_FILIAL " _
        & " as lancamento_filial_filial_nome , F_3915621685.F_1907459465  as F_1907459465 , F_3915621685.lancamento_lancamento_OBS " _
        & " as lancamento_lancamento_OBS , F_3915621685.F_2282248934  as F_2282248934 , F_3915621685.F_3518719985  as F_3518719985" _
        & " , coalesce ( F_3915621685.sum_lancamento_VALOR_INICIAL_ , 0 )  as sum_lancamento_VALOR_INICIAL_ , cast" _
        & " ( null as double precision )  as F_2714866558 from TMP_DESPESA_3 F_3915621685 inner join lancamentos F_3233442989" _
        & " on ( ( ( F_3915621685.F_3518719985 = F_3233442989.LANCAMENTO and F_3915621685.lancamento_lancamento_OBS =" _
        & " F_3233442989.OBS ) ) ) inner join Filiais F_2463272603 on ( ( ( F_3915621685.F_2282248934 = F_2463272603.FILIAL" _
        & " ) ) )  union all  select DISTINCT F_2463272603.COD_FILIAL  as lancamento_filial_filial_nome , F_1308918612.F_1907459465" _
        & "  as F_1907459465 , F_1308918612.lancamento_lancamento_OBS  as lancamento_lancamento_OBS ," _
        & " F_1308918612.F_2282248934  as F_2282248934 , F_1308918612.F_3518719985  as F_3518719985 ," _
        & " cast ( null as double precision )  as sum_lancamento_VALOR_INICIAL_ , coalesce ( F_1308918612.F_2714866558 , 0 )" _
        & "  as F_2714866558 from TMP_DESPESA_2 F_1308918612 inner join lancamentos F_3233442989 on ( ( ( F_1308918612.F_3518719985" _
        & " = F_3233442989.LANCAMENTO and F_1308918612.lancamento_lancamento_OBS = F_3233442989.OBS ) ) ) inner join" _
        & " Filiais F_2463272603 on ( ( ( F_1308918612.F_2282248934 = F_2463272603.FILIAL ) ) )   ) v "
    wSQLfim = " group by v.lancamento_filial_filial_nome , v.F_1907459465 , v.lancamento_lancamento_OBS , v.F_2282248934 , v.F_3518719985;"
    
    'wSQL = wSQLini & wSQLfim
    'Set adoDespesas.Recordset = de.cncMwts.Execute(wSQL).Clone
    'de.cncMwts.Close
    
    Call txt2list("DespesasCriterios.txt", lstCriterios)
    
    cmdPesq_Click
   
End Sub

Private Sub grid_Despesas_DblClick()
Dim vetorCriterio As Variant
Dim valorAnterior As String
    'vetorCriterio = Split(adoDespesas.Recordset("LANCAMENTO_LANCAMENTO_OBS").value, " ")
    'txtCriterio = vetorCriterio(0) & " " & vetorCriterio(1)
    'valorAnterior = adoDespesas.Recordset
    txtCriterio = adoDespesas.Recordset("LANCAMENTO_LANCAMENTO_OBS").value
End Sub

Private Sub grid_Despesas_L_DblClick()
Dim vetorCriterio As Variant
    'vetorCriterio = Split(adoDespesas.Recordset("LANCAMENTO_LANCAMENTO_OBS").value, " ")
    'txtCriterio = vetorCriterio(0) & " " & vetorCriterio(1)
    txtCriterio = adoDespesas.Recordset("LANCAMENTO_LANCAMENTO_OBS").value
End Sub

Private Sub grid_Off_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then grid_Off_LostFocus
End Sub

Private Sub grid_Off_LostFocus()
    grid_Off.Visible = False
End Sub

Private Sub lbCount_Click()
    grid_Off.Visible = True
    grid_Off.SetFocus
End Sub

Private Sub TXT_LOGO_Change()
   TXT_LOGO2.BoundText = TXT_LOGO.BoundText
   cmdPesq_Click
End Sub

Private Sub TXT_LOGO_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then
        Sendkeys "{tab}"
        cmdPesq_Click
      End If
End Sub

Private Sub TXT_LOGO2_Change()
    TXT_LOGO.BoundText = TXT_LOGO2.BoundText
End Sub

Private Sub txtCriterio_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

