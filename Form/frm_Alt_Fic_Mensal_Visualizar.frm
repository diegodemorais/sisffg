VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_Alt_Fic_Mensal_Visualizar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VISUALIZAÇÃO COMPLETA"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11730
   Icon            =   "frm_Alt_Fic_Mensal_Visualizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11730
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TXT_CAMPOS 
      Alignment       =   2  'Center
      DataField       =   "M_DT_DEM"
      DataSource      =   "ADOREG"
      Height          =   285
      Index           =   8
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   1110
      Width           =   1200
   End
   Begin VB.TextBox TXT_CAMPOS 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      DataField       =   "M_F_COD"
      DataSource      =   "ADOREG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   7
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   1470
      Width           =   855
   End
   Begin MSAdodcLib.Adodc adoConta 
      Height          =   330
      Left            =   4680
      Top             =   7680
      Visible         =   0   'False
      Width           =   2970
      _ExtentX        =   5239
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
   Begin MSDataGridLib.DataGrid Grid 
      Bindings        =   "frm_Alt_Fic_Mensal_Visualizar.frx":12D2
      Height          =   6300
      Left            =   8220
      TabIndex        =   0
      Top             =   825
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   11113
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
      Caption         =   "FICHAS"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "F_NOME"
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
      BeginProperty Column01 
         DataField       =   "M_NFICHA"
         Caption         =   "Nº FICHA"
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
            ColumnWidth     =   2129,953
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   780,095
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt_SaldoAnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008080FF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   7170
      Width           =   1470
   End
   Begin MSDataGridLib.DataGrid grid_Conta 
      Bindings        =   "frm_Alt_Fic_Mensal_Visualizar.frx":12E7
      Height          =   3435
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   6059
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
      Caption         =   "CONTAS"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "C_DT"
         Caption         =   "DATA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "TP_DESC"
         Caption         =   "CONTAS"
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
         DataField       =   "C_DESC"
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
         DataField       =   "C_VALOR"
         Caption         =   "VALOR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "R$ #0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "c_tp_op"
         Caption         =   "OP."
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
         DataField       =   "c_Visto"
         Caption         =   "Visto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "ok"
            FalseValue      =   "não"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "C_CODIGO"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   2129,953
         EndProperty
         BeginProperty Column02 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   2459,906
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   1170,142
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   420,095
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column06 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   0
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt_Cred 
      Alignment       =   2  'Center
      DataField       =   "Cred"
      DataSource      =   "ADOREG"
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   2070
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox TXT_TOTAL 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "R$ #.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
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
      Left            =   6510
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   7170
      Width           =   1620
   End
   Begin VB.TextBox TXT_MENOS 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "R$ #.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   7170
      Width           =   1620
   End
   Begin VB.TextBox TXT_MAIS 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "R$ #.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   7170
      Width           =   1620
   End
   Begin VB.TextBox TXT_CAMPOS 
      Alignment       =   2  'Center
      DataField       =   "M_DT_REG"
      DataSource      =   "ADOREG"
      Height          =   285
      Index           =   3
      Left            =   3390
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1110
      Width           =   1200
   End
   Begin VB.TextBox TXT_CAMPOS 
      Alignment       =   2  'Center
      DataField       =   "M_DT_ADM"
      DataSource      =   "ADOREG"
      Height          =   285
      Index           =   2
      Left            =   2085
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1110
      Width           =   1200
   End
   Begin VB.TextBox TXT_CAMPOS 
      DataField       =   "M_ANOTACAO"
      DataSource      =   "ADOREG"
      Height          =   525
      Index           =   6
      Left            =   1170
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3030
      Width           =   6975
   End
   Begin VB.TextBox TXT_CAMPOS 
      DataField       =   "M_OBS"
      DataSource      =   "ADOREG"
      Height          =   525
      Index           =   5
      Left            =   1170
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2430
      Width           =   6975
   End
   Begin VB.TextBox TXT_CAMPOS 
      DataField       =   "M_FERIAS"
      DataSource      =   "ADOREG"
      Height          =   525
      Index           =   4
      Left            =   1170
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1830
      Width           =   6975
   End
   Begin VB.TextBox TXT_ANO 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      DataField       =   "M_ANO"
      DataSource      =   "ADOREG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7470
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   660
   End
   Begin VB.TextBox txtM_MES 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      DataField       =   "data"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MMMM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      DataSource      =   "ADOREG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   1620
   End
   Begin VB.TextBox TXT_CAMPOS 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      DataField       =   "Logo"
      DataSource      =   "ADOREG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1110
      Width           =   855
   End
   Begin VB.TextBox TXT_CAMPOS 
      BackColor       =   &H00C0FFC0&
      DataField       =   "F_NOME"
      DataSource      =   "ADOREG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1470
      Width           =   6120
   End
   Begin VB.TextBox txtM_NFICHA 
      Alignment       =   1  'Right Justify
      DataField       =   "M_NFICHA"
      DataSource      =   "ADOREG"
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1110
      Width           =   900
   End
   Begin MSAdodcLib.Adodc ADOREG 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   7620
      Width           =   11730
      _ExtentX        =   20690
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
      Left            =   600
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_Visualizar.frx":12FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_Visualizar.frx":1618
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_Visualizar.frx":1932
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_Visualizar.frx":1C4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_Visualizar.frx":1F66
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_Visualizar.frx":2280
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_Visualizar.frx":259A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_Visualizar.frx":29EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   1429
      ButtonWidth     =   1376
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
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
            Caption         =   "&Editar"
            Key             =   "editar"
            Object.ToolTipText     =   "Editar Alteração (Alt+E)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Fil&trar"
            Key             =   "filtrar"
            Object.ToolTipText     =   "Filtrar (Alt+T)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   7
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Key             =   "imprimir"
            Object.ToolTipText     =   "Imprime a Ficha (Alt+I)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "(F5)"
            Key             =   "dupla"
            Object.ToolTipText     =   "Visualização Dupla de Fichas (F5)"
            ImageIndex      =   8
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         Caption         =   "Pesquisar :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   3600
         TabIndex        =   33
         Top             =   0
         Width           =   8055
         Begin VB.CommandButton cmdFiltrar 
            Height          =   555
            Left            =   7320
            Picture         =   "frm_Alt_Fic_Mensal_Visualizar.frx":2D06
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton Op 
            Caption         =   "Nº"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   135
            TabIndex        =   43
            Top             =   180
            Width           =   975
         End
         Begin VB.OptionButton Op 
            Caption         =   "(B)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   1
            Left            =   135
            TabIndex        =   42
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton Op 
            Caption         =   "Mês / Ano"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   2
            Left            =   1095
            TabIndex        =   41
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Op 
            Caption         =   "Emp."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   3
            Left            =   1095
            TabIndex        =   40
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Op 
            Caption         =   "Remover Filtro"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   2205
            TabIndex        =   39
            Top             =   420
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Frame p_MA 
            Caption         =   "Mês / Ano"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   3780
            TabIndex        =   37
            Top             =   105
            Visible         =   0   'False
            Width           =   2295
            Begin VB.ComboBox txt_PMes 
               Height          =   315
               ItemData        =   "frm_Alt_Fic_Mensal_Visualizar.frx":3010
               Left            =   195
               List            =   "frm_Alt_Fic_Mensal_Visualizar.frx":3038
               TabIndex        =   36
               Top             =   210
               Width           =   570
            End
            Begin VB.TextBox txt_PAno 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   1095
               TabIndex        =   38
               Top             =   210
               Width           =   615
            End
            Begin VB.Line Line1 
               BorderWidth     =   3
               X1              =   840
               X2              =   1020
               Y1              =   480
               Y2              =   255
            End
         End
         Begin VB.Frame p_Dg 
            Caption         =   "Digite :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   3810
            TabIndex        =   34
            Top             =   120
            Visible         =   0   'False
            Width           =   3495
            Begin VB.TextBox txt_Pesq 
               Height          =   285
               Left            =   240
               TabIndex        =   35
               Top             =   210
               Width           =   3015
            End
         End
      End
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "(D)"
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
      Index           =   13
      Left            =   5175
      TabIndex        =   46
      Top             =   870
      Width           =   270
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Devedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   8355
      TabIndex        =   31
      Top             =   7215
      Width           =   1785
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "=      TOTAL "
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
      Left            =   5265
      TabIndex        =   28
      Top             =   7245
      Width           =   1260
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   11
      Left            =   2640
      TabIndex        =   27
      Top             =   7080
      Width           =   150
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   10
      Left            =   165
      TabIndex        =   26
      Top             =   7110
      Width           =   240
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "®"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   3885
      TabIndex        =   25
      Top             =   855
      Width           =   210
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   2565
      TabIndex        =   24
      Top             =   840
      Width           =   285
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ANOTAÇÃO"
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
      Index           =   7
      Left            =   120
      TabIndex        =   23
      Top             =   3030
      Width           =   1020
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "OBS"
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
      Index           =   6
      Left            =   120
      TabIndex        =   22
      Top             =   2430
      Width           =   390
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "(F)"
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
      Index           =   5
      Left            =   120
      TabIndex        =   21
      Top             =   1830
      Width           =   240
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ANO"
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
      Index           =   4
      Left            =   7605
      TabIndex        =   20
      Top             =   870
      Width           =   450
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MÊS"
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
      Left            =   5865
      TabIndex        =   19
      Top             =   870
      Width           =   1635
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(B)"
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
      Left            =   1185
      TabIndex        =   18
      Top             =   870
      Width           =   255
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "NOME"
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
      Left            =   120
      TabIndex        =   17
      Top             =   1500
      Width           =   555
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº FICHA"
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
      Left            =   120
      TabIndex        =   16
      Top             =   870
      Width           =   825
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      Height          =   450
      Left            =   8220
      Top             =   7125
      Width           =   3495
   End
End
Attribute VB_Name = "frm_Alt_Fic_Mensal_Visualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W_LD_FILTRO As Boolean
Dim W_POS As Long
Dim w_Move As Boolean
Dim W_Pv  As Boolean ' SE É A PRIMEIRA VEZ QUE ENTREA NA TELA
Dim W_FILTRO As String
Dim W_INDEX As Byte
Dim w_F5 As Boolean

Private Sub Total()
Dim ADO_TOTAL As ADODB.Recordset

On Error GoTo err1
    
    TXT_MAIS = 0
    TXT_MENOS = 0
    TXT_TOTAL = 0
    
    Set ADO_TOTAL = adoConta.Recordset.Clone
    
    If Not ADO_TOTAL.EOF Then
        ADO_TOTAL.MoveFirst
        Do While Not ADO_TOTAL.EOF
            If ADO_TOTAL.Fields("C_valor") >= 0 And ADO_TOTAL.Fields("C_Tp_OP") = "+" Then
                TXT_MAIS = CDbl(TXT_MAIS) + ADO_TOTAL.Fields("C_VALOR")
            ElseIf ADO_TOTAL.Fields("C_valor") < 0 And ADO_TOTAL.Fields("C_Tp_OP") = "-" Then
                TXT_MENOS = CDbl(TXT_MENOS) + ADO_TOTAL.Fields("C_VALOR")
            End If
            ADO_TOTAL.MoveNext
        Loop
        
        TXT_TOTAL = CDbl(TXT_MAIS) - CDbl(TXT_MENOS)
    End If
    
    TXT_TOTAL = Format(CDbl(TXT_MENOS) + CDbl(TXT_MAIS), "R$ 0.00")
    TXT_MAIS = Format(TXT_MAIS, "R$ #0.00")
    TXT_MENOS = Format(TXT_MENOS, "R$ #0.00")
    
    
    'muda cor do total
    If TXT_TOTAL < 0 Then
        TXT_TOTAL.ForeColor = vbRed
    Else
        TXT_TOTAL.ForeColor = vbBlue
    End If
    
    
    w_saldo = de.cnc.Execute("Select F_SALDO_ANT FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
    'Saldo restante da ficha
    txt_SaldoAnt = IIf(IsNull(w_saldo), 0, w_saldo)
    If txt_SaldoAnt < 0 Then
        txt_SaldoAnt.ForeColor = vbRed
    Else
        txt_SaldoAnt.ForeColor = vbBlue
    End If
    txt_SaldoAnt = Format(txt_SaldoAnt, "R$ 0.00")


sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub




Private Sub cmdFiltrar_Click()
    
    FILTRAR W_INDEX


  '  Pause 0.5
    
  '  p_Pesq.Visible = False
End Sub

Private Sub Form_Activate()

    On Error Resume Next
    If Not W_FILTRO = "" And Not W_FILTRO = "0" Then W_FILTRO = ADOREG.Recordset.Filter
    
On Error GoTo err1
    
    'If UCase(frmLogin.txtUserName) = UCase(NomeMestre) Then
    BarraF.Buttons("dupla").Visible = True
    
    de.rsTAB_FICHA_MENS.Requery
    
    If de.rscmdSqlVisualizarFichas.State = 1 Then de.rscmdSqlVisualizarFichas.Close
    de.cmdSqlVisualizarFichas txt_PAno, txt_PMes
    Set ADOREG.Recordset = de.rscmdSqlVisualizarFichas.Clone          'de.cnc.Execute("SELECT TAB_FICHA_MENS.M_NFICHA, TAB_FUNCIONARIO.F_Cod_L AS LOGO, TAB_FICHA_MENS.M_ANO, TAB_FUNCIONARIO.F_NOME, TAB_FUNCIONARIO.F_DT_ADM, TAB_FUNCIONARIO.F_DT_REG, TAB_FICHA_MENS.M_FERIAS, TAB_FICHA_MENS.M_OBS, TAB_FUNCIONARIO.F_ANOTACAO, TAB_FICHA_MENS.M_TOTAL_MAIS AS MAIS, TAB_FICHA_MENS.M_TOTAL_MENOS AS MENOS, TAB_FICHA_MENS.M_TOTAL_MAIS - TAB_FICHA_MENS.M_TOTAL_MENOS AS TOTAL, '01/' + str(TAB_FICHA_MENS.M_MES) + '/' + str(TAB_FICHA_MENS.M_ANO) AS data, TAB_FICHA_MENS.M_MES as M_MES, TAB_FICHA_MENS.m_bloq as Bloq, Tab_Funcionario.F_Cod_Cred as Cred FROM TAB_FICHA_MENS, TAB_FUNCIONARIO WHERE TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo  Order By  TAB_FICHA_MENS.M_MES, TAB_FUNCIONARIO.F_Nome ").Clone
    de.rscmdSqlVisualizarFichas.Close
 
    If de.rscmdSqlVisualizarFichas.State = 1 Then de.rscmdSqlVisualizarFichas.Close

    de.cmdSqlVisualizarFichas txt_PAno, txt_PMes
    Set ADOREG.Recordset = de.rscmdSqlVisualizarFichas.Clone          'de.cnc.Execute("SELECT TAB_FICHA_MENS.M_NFICHA, TAB_FUNCIONARIO.F_Cod_L AS LOGO, TAB_FICHA_MENS.M_ANO, TAB_FUNCIONARIO.F_NOME, TAB_FUNCIONARIO.F_DT_ADM, TAB_FUNCIONARIO.F_DT_REG, TAB_FICHA_MENS.M_FERIAS, TAB_FICHA_MENS.M_OBS, TAB_FUNCIONARIO.F_ANOTACAO, TAB_FICHA_MENS.M_TOTAL_MAIS AS MAIS, TAB_FICHA_MENS.M_TOTAL_MENOS AS MENOS, TAB_FICHA_MENS.M_TOTAL_MAIS - TAB_FICHA_MENS.M_TOTAL_MENOS AS TOTAL, '01/' + str(TAB_FICHA_MENS.M_MES) + '/' + str(TAB_FICHA_MENS.M_ANO) AS data, TAB_FICHA_MENS.M_MES as M_MES, TAB_FICHA_MENS.m_bloq as Bloq, Tab_Funcionario.F_Cod_Cred as Cred FROM TAB_FICHA_MENS, TAB_FUNCIONARIO WHERE TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo  Order By  TAB_FICHA_MENS.M_MES, TAB_FUNCIONARIO.F_Nome ").Clone
 
 
    
    If Not W_FILTRO = "" And Not W_FILTRO = "0" Then ADOREG.Recordset.Filter = W_FILTRO
    
    If W_POS <> 0 Then ADOREG.Recordset.Find "m_nficha = " & W_POS
    
    If ADOREG.Recordset.EOF Then ADOREG.Recordset.Filter = 0  ' ADOREG.Recordset.MoveLast
    Pause 0.5
    w_F5 = False
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


'*** Caption no navegador ***

Private Sub ADOREG_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1
    
 If Not ADOREG.Recordset.EOF Then
    grid_Conta.Visible = True
    
    ADOREG.Caption = "REGISTRO : " & ADOREG.Recordset.AbsolutePosition & " / " & ADOREG.Recordset.RecordCount & IIf(W_LD_FILTRO = True, " (FILTRADO)", "")
         
    Set adoConta.Recordset = de.cnc.Execute("SELECT TAB_DESC_CALC.C_CODIGO, TAB_DESC_CALC.C_N_FICHA, TAB_DESC_CALC.C_DT, TAB_TP_CONTA.TP_DESC, TAB_DESC_CALC.C_TP_OP, TAB_DESC_CALC.C_VALOR, TAB_DESC_CALC.C_VISTO, TAB_DESC_CALC.C_DESC FROM TAB_DESC_CALC, TAB_TP_CONTA WHERE TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_DESC_CALC.C_N_FICHA = " & ADOREG.Recordset.Fields("M_Nficha") & ") ORDER BY TAB_TP_CONTA.TP_DESC, TAB_DESC_CALC.C_TP_OP DESC").Clone
    adoConta.Refresh
   ' Pause 0.3
    Total
    TXT_TOTAL.Refresh
    
    
    
    
Else
    grid_Conta.Visible = False
End If

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
        Case "imprimir": Imprimir
        Case "dupla": VisDupla
        
    End Select
End Sub


'*** Rotinas ***
Sub VisDupla()
    
    frm_Alt_Fic_Mensal_Visualizar_Dupla.Show 1
    
End Sub

Sub Imprimir()
 
On Error GoTo err1
    
    FRM_IMP_F.TXT_LOGO = TXT_CAMPOS(0)
    FRM_IMP_F.TXT_MES = Format(ADOREG.Recordset.Fields("DATA"), "MM")
    FRM_IMP_F.TXT_ANO = TXT_ANO
    FRM_IMP_F.dbNome = TXT_CAMPOS(1)

    FRM_IMP_F.Show 1
    
    If FRM_IMP_F.txt_State = "F" Then
       MsgBox "Impressão Cancelada!", vbCritical
    Else
        If de.rscmdRelFichaMensal.State = 1 Then de.rscmdRelFichaMensal.Close
        If de.rscmdRelFichaMensal_TRIPA.State = 1 Then de.rscmdRelFichaMensal_TRIPA.Close
        
        de.cmdRelFichaMensal FRM_IMP_F.TXT_MES, FRM_IMP_F.TXT_ANO, FRM_IMP_F.dbNome, FRM_IMP_F.TXT_LOGO
        de.cmdRelFichaMensal_TRIPA FRM_IMP_F.TXT_MES, FRM_IMP_F.TXT_ANO, FRM_IMP_F.dbNome & "%", FRM_IMP_F.TXT_LOGO & "%"
        
            Set AdoItem1 = de.rscmdRelFichaMensal_TRIPA.Fields(6).Value
            Criar_RPT_TRIPA de.rscmdRelFichaMensal_TRIPA, AdoItem1
        
        
        rptFichaMensal.Show 1
'        rptFichaMensal_Tripa.PrintReport False, rptRangeAllPages
    End If
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume sair
    
End Sub







Private Sub Editar()

On Error GoTo Prox

If ADOREG.Recordset.Fields("BLOQ") = False Then
    
    W_POS = txtM_NFICHA
    
    If Not frm_Alt_Fic_Mensal.ADOREG.Recordset.State Then
    '    frm_Alt_Fic_Mensal.ADO_GRID.Recordset.Filter = "m_nficha = " & txtM_NFICHA
    '    frm_Alt_Fic_Mensal.ADOREG.Recordset.Filter = "m_nficha = " & txtM_NFICHA
    End If
    
Prox:
On Error GoTo err1
    frm_Alt_Fic_Mensal.Show 1
    
Else
    MsgBox "Não é possível alterar uma ficha do mês retrasado ao atual!", vbExclamation
End If

    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub Fechar()
On Error GoTo err1
    If de.rsTAB_FICHA_MENS.State = 1 Then
        de.rsTAB_DESC_CALC.Requery
        de.rsTAB_FICHA_MENS.Requery
    End If
    
    W_FILTRO = 0
    ADOREG.Recordset.Filter = 0
    Unload Me
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub FILTRAR(Index As Byte)
Dim w_resp As String
Dim W_CAMPO As String


On Error GoTo err1
    
    w_resp = Index + 1
    
    
    If Not w_resp = "" And IsNumeric(w_resp) And w_resp >= 1 And w_resp <= 5 Then
        Select Case w_resp
        'Nº
        Case 1:
            w_resp = "Nº"
            W_CAMPO = "M_F_Cod"
        'LOGO
        Case 2:
            w_resp = "LOGO"
            W_CAMPO = "LOGO"
        'MÊS/ANO
        Case 3:
            w_resp = "MÊS / ANO"
            W_CAMPO = "M_MES"
            W_CAMPO2 = "M_ANO"
        'EMP
        Case 4:
            w_resp = "EMP."
            W_CAMPO = "F_NOME"
        
        
        '*** REMOVE O FILTRO ****
        Case 5:
            If Not ADOREG.Recordset.Filter = 0 Then
                W_LD_FILTRO = False
                ADOREG.Recordset.Filter = 0
                Set ADOREG.Recordset = de.rscmdSqlVisualizarFichas.Clone

            End If
        End Select
        
        If Not w_resp = "5" Then
            If w_resp = "Nº" Then
                W_FILTRO = W_CAMPO & " = " & txt_Pesq
                W_LD_FILTRO = True
                ADOREG.Recordset.Filter = W_FILTRO
            
            ElseIf w_resp = "LOGO" Or w_resp = "EMP." Then
                W_FILTRO = W_CAMPO & " LIKE '%" & txt_Pesq & "%'"
                W_LD_FILTRO = True
                ADOREG.Recordset.Filter = W_FILTRO
            
            Else
                W_FILTRO = txt_PMes
                W_FILTRO1 = txt_PAno
                
                If Not W_FILTRO = "" And IsNumeric(W_FILTRO) And IsNumeric(W_FILTRO1) And Len(W_FILTRO1) = 4 Then
                    If de.rscmdSqlVisualizarFichas.State = 1 Then de.rscmdSqlVisualizarFichas.Close
                    de.cmdSqlVisualizarFichas W_FILTRO1, W_FILTRO
                    Set ADOREG.Recordset = de.rscmdSqlVisualizarFichas.Clone          'de.cnc.Execute("SELECT TAB_FICHA_MENS.M_NFICHA, TAB_FUNCIONARIO.F_Cod_L AS LOGO, TAB_FICHA_MENS.M_ANO, TAB_FUNCIONARIO.F_NOME, TAB_FUNCIONARIO.F_DT_ADM, TAB_FUNCIONARIO.F_DT_REG, TAB_FICHA_MENS.M_FERIAS, TAB_FICHA_MENS.M_OBS, TAB_FUNCIONARIO.F_ANOTACAO, TAB_FICHA_MENS.M_TOTAL_MAIS AS MAIS, TAB_FICHA_MENS.M_TOTAL_MENOS AS MENOS, TAB_FICHA_MENS.M_TOTAL_MAIS - TAB_FICHA_MENS.M_TOTAL_MENOS AS TOTAL, '01/' + str(TAB_FICHA_MENS.M_MES) + '/' + str(TAB_FICHA_MENS.M_ANO) AS data, TAB_FICHA_MENS.M_MES as M_MES, TAB_FICHA_MENS.m_bloq as Bloq, Tab_Funcionario.F_Cod_Cred as Cred FROM TAB_FICHA_MENS, TAB_FUNCIONARIO WHERE TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo  Order By  TAB_FICHA_MENS.M_MES, TAB_FUNCIONARIO.F_Nome ").Clone
                    W_LD_FILTRO = True
                End If
                                   
            End If
        End If
        If ADOREG.Recordset.RecordCount <= 0 Then
            MsgBox "Não existe ficha com a descrição solicitada!", vbExclamation
                W_LD_FILTRO = False
                ADOREG.Recordset.Filter = 0
                Set ADOREG.Recordset = de.rscmdSqlVisualizarFichas.Clone
        End If
            
    End If
    
    
sair:
    Exit Sub
err1:
    If Err.Number = 3001 Then
       ' MsgBox "Dados inválidos para Filtragem!", vbCritical
    ElseIf Err.Number <> 13 Then
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    End If
        W_LD_FILTRO = False
        Resume sair

End Sub











Private Sub grid_Conta_DblClick()
    
If ADOREG.Recordset.Fields("BLOQ") = 0 Then
    
    W_POS = txtM_NFICHA
    frm_Alt_Desc_Calc.lb_form = "visualizar"
    frm_Alt_Desc_Calc.TXT_NFICHA_CAD = txtM_NFICHA
    frm_Alt_Desc_Calc.LB_FUNC.Caption = TXT_CAMPOS(1).Text
    frm_Alt_Desc_Calc.Show 1

Else
    MsgBox "Não é possível alterar uma ficha anterior ao mês passado!", vbExclamation
End If

End Sub


Private Sub Grid_DblClick()
    Editar
End Sub

'--------- Ao Pressionar uma Tecla -----------

Private Sub grid_Conta_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 69 Then
        
        grid_Conta_DblClick
        
    Else
        Keys KeyCode, Shift
    End If
End Sub
Private Sub GRID_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub









Private Sub Op_Click(Index As Integer)
  If Index = 2 Then
     p_Dg.Visible = False
     p_MA.Visible = True
     txt_PMes.SetFocus
  ElseIf Index = 4 Then
     p_Dg.Visible = False
     p_MA.Visible = False
     
     W_LD_FILTRO = False
     ADOREG.Recordset.Filter = 0
     Set ADOREG.Recordset = de.rscmdSqlVisualizarFichas.Clone
  Else
     p_Dg.Visible = True
     p_MA.Visible = False
     txt_Pesq.SetFocus
  End If

  W_INDEX = Index
End Sub

Private Sub TXT_CAMPOS_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub TXT_MAIS_GotFocus()
    Grid.SetFocus
End Sub
Private Sub TXT_MENOS_GotFocus()
    Grid.SetFocus
End Sub




Private Sub txt_PAno_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
    If KeyCode = 13 Then
        cmdFiltrar_Click
        Grid.SetFocus
    End If
End Sub

Private Sub txt_Pesq_Change()
    If Op(0).Value = False Then
        
        cmdFiltrar_Click
        
    End If
End Sub

Private Sub txt_Pesq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Grid.SetFocus
        cmdFiltrar_Click
    End If
    Keys KeyCode, Shift
End Sub



Private Sub txt_PMes_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
    If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub TXT_total_GotFocus()
    Grid.SetFocus
End Sub


Private Sub TXT_TOTAL_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_MAIS_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_MENOS_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub txt_ANO_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub txtM_MES_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub txtM_NFICHA_KeyDown(KeyCode As Integer, Shift As Integer)
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
    Case 84: ' "T"
            FILTRAR W_INDEX
    Case 73: ' "I"
            Imprimir
    End Select
ElseIf KeyCode = 116 And Shift = 0 And w_F5 = False Then
     If BarraF.Buttons("dupla").Visible = True Then
        w_F5 = True
        VisDupla
     End If
End If
End Sub





