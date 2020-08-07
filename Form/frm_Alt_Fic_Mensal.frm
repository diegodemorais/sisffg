VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "msCOMCTL.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form frm_Alt_Fic_Mensal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALTERAÇÃO DE FICHA MENSAL"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   Icon            =   "frm_Alt_Fic_Mensal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   9750
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "M_DT_DEM"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2700
      MaxLength       =   10
      TabIndex        =   33
      Top             =   1830
      Width           =   1095
   End
   Begin VB.TextBox txt_SaldoEmp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   360
      Left            =   8460
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   6810
      Width           =   1170
   End
   Begin VB.TextBox txt_DT_ADM 
      Alignment       =   2  'Center
      DataField       =   "M_DT_ADM"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      MaxLength       =   10
      TabIndex        =   24
      Top             =   1830
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "M_F_COD"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   29
      Top             =   1230
      Width           =   1260
   End
   Begin VB.TextBox TXT_DT_REG 
      Alignment       =   2  'Center
      DataField       =   "M_DT_REG"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   25
      Top             =   1830
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc ADO_GRID 
      Height          =   330
      Left            =   6960
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   8460
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7155
      Width           =   1170
   End
   Begin VB.TextBox TXT_OBS 
      DataField       =   "M_OBS"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3960
      Width           =   5055
   End
   Begin VB.TextBox TXT_TOTAL 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      DataField       =   "M_TOTAL"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      DataSource      =   "ADOREG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5745
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7125
      Width           =   1290
   End
   Begin MSAdodcLib.Adodc ADO_LANC 
      Height          =   330
      Left            =   600
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSDataGridLib.DataGrid GRID_L 
      Bindings        =   "frm_Alt_Fic_Mensal.frx":1042
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   -2147483639
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   13
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
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "LANÇAMENTOS"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "DATA"
         Caption         =   "DATA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "d/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "CONTA"
         Caption         =   "CONTA"
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
         DataField       =   "VALOR"
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
      BeginProperty Column03 
         DataField       =   "OP"
         Caption         =   "OP"
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
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TXT_ANO 
      Alignment       =   2  'Center
      DataField       =   "M_ANO"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4605
      TabIndex        =   3
      Top             =   1230
      Width           =   810
   End
   Begin VB.ComboBox TXT_MES 
      DataField       =   "M_MES"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frm_Alt_Fic_Mensal.frx":1059
      Left            =   3840
      List            =   "frm_Alt_Fic_Mensal.frx":1081
      TabIndex        =   2
      Top             =   1230
      Width           =   780
   End
   Begin VB.TextBox TXT_NFICHA 
      Alignment       =   2  'Center
      DataField       =   "M_NFICHA"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3840
      TabIndex        =   1
      Top             =   1830
      Width           =   1560
   End
   Begin VB.TextBox TXT_FERIAS 
      DataField       =   "M_FERIAS"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3000
      Width           =   5055
   End
   Begin VB.TextBox TXT_ANOTACAO 
      DataField       =   "M_ANOTACAO"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4860
      Width           =   5055
   End
   Begin MSDataGridLib.DataGrid Grid 
      Bindings        =   "frm_Alt_Fic_Mensal.frx":10AC
      Height          =   5775
      Left            =   5760
      TabIndex        =   0
      Top             =   960
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   10186
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   -2147483639
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "FICHA MENSAL"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "M_NFICHA"
         Caption         =   "NUM."
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
         DataField       =   "F_NOME"
         Caption         =   "EMP"
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
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADOREG 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   7575
      Width           =   9750
      _ExtentX        =   17198
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
      Left            =   6120
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
            Picture         =   "frm_Alt_Fic_Mensal.frx":10C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal.frx":13DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal.frx":16F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal.frx":1A11
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal.frx":1D2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal.frx":2045
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal.frx":235F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   1482
      ButtonWidth     =   1667
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
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
            Enabled         =   0   'False
            Caption         =   "&Salvar"
            Key             =   "salvar"
            Object.ToolTipText     =   "Salvar Alteração (Alt+S)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Cancelar"
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar Alteração (Alt+C)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E&xcluir"
            Key             =   "excluir"
            Object.ToolTipText     =   "Excluir registro (Alt+X)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "filtrar"
            Object.ToolTipText     =   "Filtrar (Alt+T)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "C&ontas"
            Key             =   "conta"
            Object.ToolTipText     =   "Incluir Contas (Alt+O)"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo TXT_FUNC 
      Bindings        =   "frm_Alt_Fic_Mensal.frx":2C39
      DataField       =   "M_F_COD"
      DataSource      =   "ADOREG"
      Height          =   315
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "F_NOME"
      BoundColumn     =   "F_Codigo"
      Text            =   ""
      Object.DataMember      =   "TAB_FUNCIONARIO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo TXT_CRED 
      Bindings        =   "frm_Alt_Fic_Mensal.frx":2C4A
      DataField       =   "M_F_COD"
      DataSource      =   "ADOREG"
      Height          =   315
      Left            =   2880
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "F_COD_CRED"
      BoundColumn     =   "F_Codigo"
      Text            =   ""
      Object.DataMember      =   "TAB_FUNCIONARIO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo txtLogo 
      Bindings        =   "frm_Alt_Fic_Mensal.frx":2C5B
      DataField       =   "M_LOGO"
      DataSource      =   "ADOREG"
      Height          =   315
      Left            =   1815
      TabIndex        =   21
      Top             =   1245
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      ListField       =   "COD_LOJ"
      BoundColumn     =   "COD_LOJ"
      Text            =   ""
      Object.DataMember      =   "TAB_L"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "(D)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2745
      TabIndex        =   34
      Top             =   1575
      Width           =   1155
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S. Empréstimo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   7125
      TabIndex        =   32
      Top             =   6855
      Width           =   1365
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   990
      Width           =   855
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "®"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1605
      TabIndex        =   28
      Top             =   1590
      Width           =   1155
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   390
      TabIndex        =   27
      Top             =   1590
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S. Devedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   7215
      TabIndex        =   26
      Top             =   7170
      Width           =   1185
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "(B)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1830
      TabIndex        =   22
      Top             =   1005
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVAÇÃO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5760
      TabIndex        =   19
      Top             =   6900
      Width           =   1305
   End
   Begin VB.Label TXT_FILTRO 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   16
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº FICHA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   1590
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "MÊS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   990
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(F)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ANOTAÇÃO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   375
      TabIndex        =   12
      Top             =   4635
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "EMP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ANO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   990
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   4650
      Left            =   120
      Top             =   960
      Width           =   5535
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      Height          =   825
      Left            =   7095
      Top             =   6750
      Width           =   2580
   End
End
Attribute VB_Name = "frm_Alt_Fic_Mensal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W_LD_FILTRO As Boolean
Dim V_MOVE As Boolean
Dim V_MOVE_GRID As Boolean
Dim W_FILTRO As String


'ADODB.Recordset
Private Sub adoReg_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1
    If Not adoReg.Recordset.EOF Then
        adoReg.Caption = "REGISTRO : " & adoReg.Recordset.AbsolutePosition & " / " & adoReg.Recordset.RecordCount & IIf(W_LD_FILTRO = True, " (FILTRADO)", "")
    
            '*** DESABILITA O EDITAR ****
            
            If adoReg.Recordset.Fields("M_BLOQ") = True Then
                 BarraF.Buttons("editar").Enabled = False
            Else
                 BarraF.Buttons("editar").Enabled = True
            End If
            
            If adoReg.Recordset.Fields("m_TOTAL") < 0 Then
                TXT_TOTAL.ForeColor = vbRed
            Else
                TXT_TOTAL.ForeColor = vbBlue
            End If
    
            If V_MOVE = True Then
                 On Error Resume Next
                 V_MOVE = False
                 'ADO_GRID.Recordset.Requery
                 If Not ADO_GRID.Recordset.EOF Then
                     
                     Select Case adReason
                        Case 12: '*** Vai p/ o Primeiro Registro ***
                            ADO_GRID.Recordset.MoveFirst
                        Case 13: '*** Vai p/ o Próximo Registro ***
                            ADO_GRID.Recordset.MoveNext
                        Case 14: '*** Vai p/ o Anterior Registro ***
                            ADO_GRID.Recordset.MovePrevious
                        Case 15: '*** Vai p/ o Ultimo Registro ***
                            ADO_GRID.Recordset.MoveLast
                     End Select
                         
                 End If
            End If
        End If
sair:
    
    V_MOVE = True
    Exit Sub

err1:
    If Not Err.Number = -2147217885 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub Form_Activate()
On Error GoTo err1

    de.rsTAB_DESC_CALC.Close
    de.TAB_DESC_CALC
    
    'Saldo DO EMPRESTIMO
    de.rsTAB_FUNCIONARIO.Requery
    w_Emprest = de.cnc.Execute("Select F_EMPRESTIMO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & adoReg.Recordset.Fields("M_F_COD") & "").Fields(0)
    txt_SaldoEmp = IIf(IsNull(w_Emprest), Format(0, "R$ 0.00"), Format(w_Emprest, "R$ 0.00"))
    
    Set ADO_LANC.Recordset = de.cnc.Execute("SELECT TAB_DESC_CALC.C_DT AS DATA, 'CT: ' + TAB_TP_CONTA.TP_DESC + '     DESC: ' + TAB_DESC_CALC.C_DESC AS CONTA, TAB_DESC_CALC.C_VALOR AS VALOR, TAB_DESC_CALC.C_TP_OP AS OP FROM TAB_DESC_CALC, TAB_TP_CONTA WHERE TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_DESC_CALC.C_N_FICHA = " & frm_Alt_Fic_Mensal_Visualizar.adoReg.Recordset.Fields("M_NFICHA") & ")").Clone
    ADO_LANC.Refresh
   
    '*** CALCULA O TOTAL - APÓS O NOVO VALOR ***
    W_MAIS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MAIS FROM TAB_DESC_CALC WHERE C_Tp_Op = '+' and C_VALOR > 0 AND (C_N_FICHA = " & frm_Alt_Fic_Mensal_Visualizar.adoReg.Recordset.Fields("M_NFICHA") & ")").Fields("MAIS")
    W_MENOS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MENOS FROM TAB_DESC_CALC WHERE C_Tp_Op = '-' and C_VALOR < 0 AND (C_N_FICHA = " & frm_Alt_Fic_Mensal_Visualizar.adoReg.Recordset.Fields("M_NFICHA") & ")").Fields("MENOS")
    
    W_TOTAL = IIf(IsNull(W_MENOS), 0, W_MENOS) + IIf(IsNull(W_MAIS), 0, W_MAIS)
 
    If adoReg.Recordset.Fields("m_TOTAL") <> CDbl(W_TOTAL) Then
        TXT_TOTAL = Format(W_TOTAL, "R$ 0.00")
        adoReg.Recordset.Fields("m_TOTAL") = TXT_TOTAL
    End If

    adoReg.Recordset.UpdateBatch adAffectCurrent
    
    If adoReg.Recordset.Fields("m_TOTAL") < 0 Then
        TXT_TOTAL.ForeColor = vbRed
    Else
        TXT_TOTAL.ForeColor = vbBlue
    End If


    Set ADO_LANC.Recordset = de.cnc.Execute("SELECT TAB_DESC_CALC.C_DT AS DATA, 'CT: ' + TAB_TP_CONTA.TP_DESC + '     DESC: ' + TAB_DESC_CALC.C_DESC AS CONTA, TAB_DESC_CALC.C_VALOR AS VALOR, TAB_DESC_CALC.C_TP_OP AS OP FROM TAB_DESC_CALC, TAB_TP_CONTA WHERE TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_DESC_CALC.C_N_FICHA = " & frm_Alt_Fic_Mensal_Visualizar.adoReg.Recordset.Fields("M_NFICHA") & ")").Clone
    ADO_LANC.Refresh
    
    txtLogo.SetFocus
    
sair:
   ' Salvar
    Exit Sub
err1:
    If Err.Number <> 3705 And Err.Number <> -2147217864 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
    Set ADO_LANC.Recordset = de.cnc.Execute("SELECT TAB_DESC_CALC.C_DT AS DATA, 'CT: ' + TAB_TP_CONTA.TP_DESC + '     DESC: ' + TAB_DESC_CALC.C_DESC AS CONTA, TAB_DESC_CALC.C_VALOR AS VALOR, TAB_DESC_CALC.C_TP_OP AS OP FROM TAB_DESC_CALC, TAB_TP_CONTA WHERE TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_DESC_CALC.C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("M_NFICHA") & ")").Clone
    Cancelar
    Cancelar

    Resume sair
End Sub

Private Sub Form_Load()
On Error GoTo err1
      
    If de.rscmdBase.State = 1 Then de.rscmdBase.Close
    de.rscmdBase.Open "SELECT TAB_FICHA_MENS.*  FROM TAB_FICHA_MENS Where m_nficha = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("M_NFICHA") & "", , adOpenStatic, adLockOptimistic
    Set adoReg.Recordset = de.rscmdBase.Clone   '.Clone

    V_MOVE = False
    Set ADO_GRID.Recordset = de.cnc.Execute("SELECT TAB_FICHA_MENS.* , TAB_FUNCIONARIO.F_NOME FROM TAB_FUNCIONARIO, TAB_FICHA_MENS WHERE TAB_FUNCIONARIO.F_CODIGO = TAB_FICHA_MENS.M_F_COD and m_nficha = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("M_NFICHA") & " Order By TAB_FICHA_MENS.M_ANO, TAB_FICHA_MENS.M_Mes").Clone
    Set ADO_LANC.Recordset = de.cnc.Execute("SELECT TAB_DESC_CALC.C_DT AS DATA, 'CT: ' + TAB_TP_CONTA.TP_DESC + '     DESC: ' + TAB_DESC_CALC.C_DESC AS CONTA, TAB_DESC_CALC.C_VALOR AS VALOR, TAB_DESC_CALC.C_TP_OP AS OP FROM TAB_DESC_CALC, TAB_TP_CONTA WHERE TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_DESC_CALC.C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("M_NFICHA") & ")").Clone
    
    If TXT_FILTRO <> "" Then adoReg.Recordset.Filter = TXT_FILTRO
    
    V_MOVE = True
    
    'Saldo restante da ficha
    W_SALDO = de.cnc.Execute("Select F_SALDO_ANT FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & adoReg.Recordset.Fields("M_F_COD") & "").Fields(0)
    txt_SaldoAnt = IIf(IsNull(W_SALDO), Format(0, "R$ 0.00"), Format(W_SALDO, "R$ 0.00"))
    'Saldo DO EMPRESTIMO
    w_Emprest = de.cnc.Execute("Select F_EMPRESTIMO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & adoReg.Recordset.Fields("M_F_COD") & "").Fields(0)
    txt_SaldoEmp = IIf(IsNull(w_Emprest), Format(0, "R$ 0.00"), Format(w_Emprest, "R$ 0.00"))

    If txt_SaldoAnt < 0 Then
        txt_SaldoAnt.ForeColor = vbRed
    Else
        txt_SaldoAnt.ForeColor = vbBlue
    End If
        
    If de.rscmdBase.State = 1 Then adoReg.Refresh
    Editar

sair:
    Exit Sub
err1:
    If Not Err.Number = 3705 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub ADO_GRID_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1

    If V_MOVE = True Then
        V_MOVE = False
        adoReg.Recordset.UpdateBatch ' adAffectCurrent
        adoReg.Refresh
        adoReg.Recordset.Requery
        adoReg.Recordset.Move ADO_GRID.Recordset.AbsolutePosition - 1
        V_MOVE = True
        Set ADO_LANC.Recordset = de.cnc.Execute("SELECT TAB_DESC_CALC.C_DT AS DATA, 'CT: ' + TAB_TP_CONTA.TP_DESC + '     DESC: ' + TAB_DESC_CALC.C_DESC AS CONTA, TAB_DESC_CALC.C_VALOR AS VALOR, TAB_DESC_CALC.C_TP_OP AS OP FROM TAB_DESC_CALC, TAB_TP_CONTA WHERE TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_DESC_CALC.C_N_FICHA = " & ADO_GRID.Recordset.Fields("M_NFICHA") & ")").Clone
    End If

sair:
    V_MOVE = True
    Exit Sub
err1:
    If Not (Err.Number = -2147217885) And Not (Err.Number = 3021) And Not (Err.Number = 91) Then MsgBox Err.Number & " : " & Err.Description, vbCritical
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
        Case "conta": CONTA
    End Select
End Sub


'*** Rotinas ***
Private Sub CONTA()
 
If adoReg.Recordset.Fields("M_BLOQ") = False Then
 
    frm_Alt_Desc_Calc.lb_form = "mensal"
    frm_Alt_Desc_Calc.LB_FUNC.Caption = TXT_FUNC.text

    frm_Alt_Desc_Calc.Show 1
    
Else
    MsgBox "Não é possível alterar uma ficha anterior ao mês passado!", vbExclamation
End If
    
End Sub

Private Sub Cancelar()
On Error GoTo err1
    W_FILTRO = adoReg.Recordset.Filter
    pos = adoReg.Recordset.Fields("m_nficha")
    adoReg.Recordset.CancelBatch adAffectCurrent
    adoReg.Refresh
    
    If W_FILTRO <> "0" Then adoReg.Recordset.Filter = W_FILTRO
    adoReg.Recordset.MoveFirst
    Do While Not adoReg.Recordset.EOF
      If adoReg.Recordset.Fields("m_nficha") = pos Then Exit Do
      adoReg.Recordset.MoveNext
    Loop

    Pause 0.3
   Editar
    
       '*** DESABILITA O EDITAR ****
   If adoReg.Recordset.Fields("M_BLOQ") = True Then
        BarraF.Buttons("editar").Enabled = False
   Else
        BarraF.Buttons("editar").Enabled = True
   End If
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub Editar()
On Error GoTo err1
    
If adoReg.Recordset.Fields("M_BLOQ") = False Then

    BarraF.Buttons("salvar").Enabled = Not BarraF.Buttons("salvar").Enabled
    BarraF.Buttons("cancelar").Enabled = Not BarraF.Buttons("cancelar").Enabled
    BarraF.Buttons("editar").Enabled = Not BarraF.Buttons("editar").Enabled
    
    Grid.Enabled = Not Grid.Enabled
    
    txtLogo.Enabled = Not txtLogo.Enabled
'    TXT_MES.Enabled = Not TXT_MES.Enabled
'    TXT_ANO.Enabled = Not TXT_ANO.Enabled
'    TXT_ANO.Enabled = Not TXT_ANO.Enabled
    txt_DT_ADM.Enabled = Not txt_DT_ADM.Enabled
    TXT_DT_REG.Enabled = Not TXT_DT_REG.Enabled


    TXT_FUNC.Enabled = Not TXT_FUNC.Enabled
    TXT_FERIAS.Enabled = Not TXT_FERIAS.Enabled
    TXT_OBS.Enabled = Not TXT_OBS.Enabled
    TXT_ANOTACAO.Enabled = Not TXT_ANOTACAO.Enabled
    If BarraF.Buttons("salvar").Enabled = False Then
        Grid.SetFocus
    Else
        'txtLogo.SetFocus
    End If

Else
    MsgBox "Não é possível alterar uma ficha anterior ao mês passado!", vbExclamation
End If
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Excluir()
On Error GoTo err1
        
        
    frm_Habilitar.Show 1
    w_PSS = frm_Habilitar.txt_Pss

If w_PSS = w_PassWordLib Then
        
    If vbYes = MsgBox("DESEJA REALMENTE EXCLUIR A FICHA MENSAL (" & TXT_NFICHA & " : " & TXT_FUNC & ")?", vbQuestion + vbYesNo) Then
        W_POS = adoReg.Recordset.AbsolutePosition - 1
        adoReg.Recordset.Delete
        adoReg.Recordset.UpdateBatch
        w_adoFiltro = adoReg.Recordset.Filter
        Form_Load
        ADO_GRID.Refresh
        adoReg.Refresh
      
        adoReg.Recordset.Filter = w_adoFiltro
        ADO_GRID.Recordset.Filter = w_adoFiltro
      
        Grid.Refresh
        w_PSS = ""
        
      '  Grid.ReBind
    End If

Else
    MsgBox "Senha de Liberação Incorreta!", vbCritical

End If
 
sair:
    Exit Sub
err1:
    If Not Err.Number = -2147467259 Then
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    Else
        MsgBox "NÃO É POSSÍVEL EXCLUIR ESTA FICHA MENSAL, DEVIDO A CÁLCULOS RELACIONADAS A ELA!", vbCritical
        adoReg.Refresh
    End If
    Resume sair
End Sub

Private Sub Fechar()
On Error GoTo err1
    If de.rsTAB_FICHA_MENS.State = 1 Then

        de.rsTAB_FICHA_MENS.Close
        de.TAB_FICHA_MENS
        
    End If
    de.rsTAB_DESC_CALC.Requery
    Unload Me
    'Hide
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub FILTRAR()
Dim w_resp As String
Dim W_CAMPO As String

Dim W_FILTRO1 As String

On Error GoTo err1
    
    w_resp = InputBox("FILTRAR PELO QUÊ ? ENTRE COM O NÚMERO DA OPÇÃO DESEJADA." & Chr(13) & Chr(13) & "1 - Nº DA FICHA" & Chr(13) & "2 - EMP" & Chr(13) & "3 - MÊS E ANO" & Chr(13) & "4 - REMOVER FILTRO *", , "1")
    
    
    If Not w_resp = "" And IsNumeric(w_resp) And w_resp >= 1 And w_resp <= 4 Then
        Select Case w_resp
        'NFICHA
        Case 1:
            w_resp = "Nº FICHA"
            W_CAMPO = "M_NFICHA"
        'FUNCIONÁRIO
        Case 2:
            w_resp = "FUNCIONÁRIO"
            W_CAMPO = "M_F_COD"
            
        'MÊS E ANO
        Case 3:
            w_resp = "MÊS E ANO"
            W_CAMPO = "M_MES"
            
        '*** REMOVE O FILTRO ****
        Case 4:
            If Not adoReg.Recordset.Filter = 0 Then
                W_LD_FILTRO = False
                adoReg.Recordset.Filter = 0
                adoReg.Refresh
            End If
        End Select
        
        If Not w_resp = "3" Then
            Select Case w_resp
            Case "Nº FICHA":
                W_FILTRO = InputBox("ENTRE COM O " & w_resp & " DESEJADO !")
            Case "FUNCIONÁRIO":
                frm_ESCOLHA_FUNC.Show 1
                W_FILTRO = frm_ESCOLHA_FUNC.TXT_FUNC_COD
            Case "MÊS E ANO":
                W_FILTRO = CDbl(InputBox("ENTRE COM O MÊS DESEJADO !", , Format(Date, "MM")))
                W_FILTRO1 = CDbl(InputBox("ENTRE COM O ANO DESEJADO !", , Format(Date, "YYYY")))
                
                If Not W_FILTRO = "" And IsNumeric(W_FILTRO) And IsNumeric(W_FILTRO1) And Len(W_FILTRO1) = 4 Then
                    W_FILTRO = "M_MES = " & W_FILTRO & " AND M_ANO = " & W_FILTRO1
                    W_LD_FILTRO = True
                    adoReg.Recordset.Filter = W_FILTRO
                    ADO_GRID.Recordset.Filter = adoReg.Recordset.Filter
                    ADO_GRID.Recordset.MoveFirst
                End If
                                
            End Select
        
                If Not W_FILTRO = "" And IsNumeric(W_FILTRO) Then
                    W_FILTRO = W_CAMPO & " = " & W_FILTRO
                    W_LD_FILTRO = True
                    adoReg.Recordset.Filter = W_FILTRO
                End If
        
        End If
    End If
    
    ADO_GRID.Refresh
    ADO_GRID.Recordset.Filter = adoReg.Recordset.Filter
    
    
sair:
    Exit Sub
err1:
    If Err.Number <> 13 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
        W_LD_FILTRO = False
        Resume sair

End Sub

Private Sub Salvar()
On Error GoTo err1
    
    adoReg.Recordset.UpdateBatch adAffectCurrent
        
    '*** Atualiza o Funcionário ****
    
    If txt_DT_ADM = "" Then
        MsgBox "Dt. ADM   não pode conter valor nulo,  portanto será inserida a data do dia!", vbCritical
        txt_DT_ADM = Date
    End If
    
        w_dt_REg = IIf(TXT_DT_REG = "", Null, Format(TXT_DT_REG, "DD/MM/YYYY"))
        
        If IsNull(w_dt_REg) Then
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_COD_L = '" & txtLogo & "', F_FERIAS = '" & TXT_FERIAS & "', F_OBS = '" & TXT_OBS & "', F_ANOTACAO = '" & TXT_ANOTACAO & "', F_DT_ADM = '" & Format(txt_DT_ADM, "DD/MM/YYYY") & "', F_DT_REG = NULL  WHERE (F_Codigo = " & TXT_FUNC.BoundText & " )", w_reg
        Else
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_COD_L = '" & txtLogo & "', F_FERIAS = '" & TXT_FERIAS & "', F_OBS = '" & TXT_OBS & "', F_ANOTACAO = '" & TXT_ANOTACAO & "', F_DT_ADM = '" & Format(txt_DT_ADM, "DD/MM/YYYY") & "', F_DT_REG = '" & w_dt_REg & "'  WHERE (F_Codigo = " & TXT_FUNC.BoundText & " )", w_reg
        End If
    
    If w_reg = 0 Then MsgBox "Não foi possível atualiza o cadastro de funcionários (as férias e observações)", vbCritical
    
    Editar
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
    
End Sub

Private Sub GRID_L_DblClick()
    CONTA
End Sub

Private Sub GRID_L_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then
        Sendkeys "{tab}"
      End If
End Sub

Private Sub TXT_ANO_GotFocus()
    Sendkeys "{home}+{end}"
End Sub

Private Sub txt_ANOTACAO_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then
        KeyCode = 0
        Sendkeys "{BACKSPACE}"
        Sendkeys "{tab}"
      End If
End Sub

Private Sub txt_DT_ADM_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub TXT_DT_REG_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub TXT_FERIAS_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then
        KeyCode = 0
        Sendkeys "{BACKSPACE}"
        Sendkeys "{tab}"
      End If
End Sub

'--------- Ao Pressionar uma Tecla -----------
Private Sub txt_FERIAS_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub TXT_FUNC_GotFocus()
    Sendkeys "{F4}"
End Sub
Private Sub txt_ANOTACAO_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub TXT_FUNC_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then
                Sendkeys "{tab}"
      End If
End Sub

Private Sub TXT_OBS_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then
        KeyCode = 0
        Sendkeys "{BACKSPACE}"
        Sendkeys "{tab}"
      End If
End Sub

Private Sub txt_OBS_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_mes_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub TXT_FUNC_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub TXT_ano_KeyUp(KeyCode As Integer, Shift As Integer)
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
           If BarraF.Buttons("editar").Enabled = True Then Editar
    Case 83: ' "S"
           If BarraF.Buttons("salvar").Enabled = True Then Salvar
    Case 67: ' "C"
           If BarraF.Buttons("cancelar").Enabled = True Then Cancelar
    Case 88: ' "X"
            Excluir
    Case 84: ' "T"
            FILTRAR
    Case 79: ' "O"
            CONTA
    End Select
End If
End Sub


Private Sub txtLogo_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub
