VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "msCOMCTL.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form frm_Alt_Funcionario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALTERAÇÃO DE EMP"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15045
   Icon            =   "frm_ALt_Funcionario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   15045
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox dtNasc 
      Alignment       =   2  'Center
      DataField       =   "F_DT_NASC"
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   57
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox ck_pg_vt 
      Caption         =   "Check1"
      DataField       =   "F_PG_VT"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   54
      Top             =   7440
      Width           =   195
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8640
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   8454143
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   8454143
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ALt_Funcionario.frx":1042
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ALt_Funcionario.frx":135C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ALt_Funcionario.frx":1676
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ALt_Funcionario.frx":1990
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ALt_Funcionario.frx":1CAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ALt_Funcionario.frx":1FC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ALt_Funcionario.frx":22DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ALt_Funcionario.frx":2BB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ALt_Funcionario.frx":48C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ALt_Funcionario.frx":4BDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ALt_Funcionario.frx":502E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ALt_Funcionario.frx":5348
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ALt_Funcionario.frx":566A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ALt_Funcionario.frx":7E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ALt_Funcionario.frx":826E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt_cod_central 
      DataField       =   "F_COD_CENTRAL"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "####"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2760
      MaxLength       =   4
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2055
   End
   Begin VB.ComboBox TXT_TIPO 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frm_ALt_Funcionario.frx":AA20
      Left            =   840
      List            =   "frm_ALt_Funcionario.frx":AA3F
      TabIndex        =   48
      Text            =   "VENDEDOR"
      Top             =   3240
      Width           =   1875
   End
   Begin VB.TextBox txtFCod 
      BackColor       =   &H80000013&
      CausesValidation=   0   'False
      DataField       =   "F_CODIGO"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdDescCalcFixo 
      Caption         =   "&PROGRAMADOS"
      Height          =   495
      Left            =   7200
      TabIndex        =   46
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txt_notas 
      DataField       =   "F_NOTAS"
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
      Height          =   2655
      Left            =   5640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   37
      Top             =   3840
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.TextBox TXT_LOJA 
      DataField       =   "F_LOJA"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      DataSource      =   "adoReg"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TXT_VR_FIXO 
      DataField       =   "F_VR_FIXO"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      DataSource      =   "adoReg"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox TXT_COMIS 
      DataField       =   "F_COMIS"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      DataSource      =   "adoReg"
      Enabled         =   0   'False
      Height          =   285
      Left            =   7320
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox TXT_VR_MINIMO 
      DataField       =   "F_VR_MINIMO"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      DataSource      =   "adoReg"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox CK_PREMIO 
      Caption         =   "Check1"
      DataField       =   "F_PREMIO"
      DataSource      =   "adoReg"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7320
      TabIndex        =   38
      Top             =   2040
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox TXT_DT_REG 
      Alignment       =   2  'Center
      DataField       =   "F_DT_REG"
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      MaxLength       =   10
      TabIndex        =   15
      Top             =   2550
      Width           =   1095
   End
   Begin VB.TextBox txt_ANOTACAO 
      DataField       =   "F_ANOTACAO"
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
      Left            =   225
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   5970
      Width           =   4875
   End
   Begin VB.TextBox TXT_DT_DEM 
      Alignment       =   2  'Center
      DataField       =   "F_DT_DEM"
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2880
      MaxLength       =   10
      TabIndex        =   13
      Top             =   2550
      Width           =   1095
   End
   Begin VB.TextBox TXT_NOME 
      DataField       =   "F_NOME"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      Height          =   315
      Left            =   240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1200
      Width           =   4815
   End
   Begin VB.TextBox TXT_FERIAS 
      DataField       =   "F_FERIAS"
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
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   3840
      Width           =   4890
   End
   Begin VB.TextBox TXT_OBS 
      DataField       =   "F_OBS"
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
      Height          =   735
      Left            =   225
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   4890
      Width           =   4890
   End
   Begin VB.TextBox txt_VPiso 
      Alignment       =   1  'Right Justify
      DataField       =   "F_VPiso"
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
      Left            =   240
      MaxLength       =   10
      TabIndex        =   9
      Top             =   6915
      Width           =   1095
   End
   Begin VB.TextBox txt_Vend 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      DataField       =   "F_Cod_CENTRAL"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "@@.@"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   8
      Top             =   7920
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.ComboBox TXT_TIPO2 
      DataField       =   "F_TIPO"
      DataSource      =   "adoReg"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frm_ALt_Funcionario.frx":AA96
      Left            =   240
      List            =   "frm_ALt_Funcionario.frx":AAB5
      TabIndex        =   7
      Text            =   "V"
      Top             =   3240
      Width           =   585
   End
   Begin VB.TextBox txt_VPiso_R 
      Alignment       =   1  'Right Justify
      DataField       =   "F_VPiso_R"
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
      Left            =   1530
      MaxLength       =   10
      TabIndex        =   6
      Top             =   6915
      Width           =   1095
   End
   Begin VB.TextBox TXT_CX_QT_VND 
      Alignment       =   1  'Right Justify
      DataField       =   "F_CX_QT_vnd"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
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
      Left            =   2820
      MaxLength       =   10
      TabIndex        =   5
      ToolTipText     =   "MÉDIA DOS QTO VENDEDORES?"
      Top             =   6915
      Width           =   1095
   End
   Begin VB.ComboBox txt_Vcto_ferias 
      DataField       =   "F_VCTO_FERIAS"
      DataSource      =   "adoReg"
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frm_ALt_Funcionario.frx":AAD4
      Left            =   4095
      List            =   "frm_ALt_Funcionario.frx":AAFC
      TabIndex        =   4
      Top             =   2550
      Width           =   690
   End
   Begin VB.CheckBox ck_pg_SFam 
      Caption         =   "Check1"
      DataField       =   "F_PG_SAL_FAM"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   7440
      Width           =   195
   End
   Begin VB.TextBox txt_NFilhos 
      Alignment       =   2  'Center
      DataField       =   "F_NUM_FILHOS"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
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
      Height          =   270
      Left            =   360
      MaxLength       =   10
      TabIndex        =   2
      Top             =   7995
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc adoReg 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   8520
      Width           =   15045
      _ExtentX        =   26538
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
   Begin MSDataGridLib.DataGrid Grid 
      Bindings        =   "frm_ALt_Funcionario.frx":AB27
      Height          =   7575
      Left            =   9600
      TabIndex        =   0
      Top             =   840
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   13361
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
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
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "F_CODIGO"
         Caption         =   "CÓDIGO"
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
         DataField       =   "CARGO"
         Caption         =   "CARGO"
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
         DataField       =   "COD_LOJ"
         Caption         =   "COD_LOJ"
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
         DataField       =   "COD_FUNC"
         Caption         =   "COD_FUNC"
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
         EndProperty
         BeginProperty Column01 
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
      EndProperty
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   1482
      ButtonWidth     =   2011
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Ativar"
            Key             =   "ativar"
            Object.ToolTipText     =   "Ativar funcionário"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Nova Ficha"
            Key             =   "nova"
            Object.Tag             =   "Nova Ficha"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fil&trar"
            Key             =   "filtrar"
            Object.ToolTipText     =   "Filtrar (Alt+T)"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo TXT_LOGO 
      Bindings        =   "frm_ALt_Funcionario.frx":AB3C
      DataField       =   "F_COD_L"
      DataSource      =   "ADOREG"
      Height          =   360
      Left            =   2040
      TabIndex        =   16
      Top             =   1800
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "COD_LOJ"
      BoundColumn     =   "COD_LOJ"
      Text            =   ""
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
   Begin MSComCtl2.DTPicker txt_DT_ADM 
      DataField       =   "f_dt_ADM"
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
      Height          =   345
      Left            =   1440
      TabIndex        =   17
      Top             =   2550
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   216465409
      CurrentDate     =   38432
   End
   Begin MSMask.MaskEdBox txtCPF 
      DataField       =   "F_CPF"
      DataSource      =   "adoReg"
      Height          =   315
      Left            =   240
      TabIndex        =   51
      Top             =   1830
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      Enabled         =   0   'False
      MaxLength       =   14
      Format          =   "###.###.###-##"
      Mask            =   "###.###.###-##"
      PromptChar      =   "_"
   End
   Begin MSDataListLib.DataCombo TXT_LOGO2 
      Bindings        =   "frm_ALt_Funcionario.frx":AB4D
      DataField       =   "F_COD_L"
      DataSource      =   "ADOREG"
      Height          =   360
      Left            =   2880
      TabIndex        =   53
      Top             =   1800
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "NUM"
      BoundColumn     =   "COD_LOJ"
      Text            =   ""
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
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Nasc.:"
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
      Left            =   3840
      TabIndex        =   56
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Paga vale transporte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   55
      Top             =   7440
      Width           =   1740
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "CPF"
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
      Left            =   240
      TabIndex        =   52
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "CÓDIGO MILLENNIUM"
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
      Left            =   2760
      TabIndex        =   50
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lbNotas 
      BackStyle       =   0  'Transparent
      Caption         =   "NOTAS"
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
      Left            =   5640
      TabIndex        =   45
      Top             =   3600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lbGerente 
      BackStyle       =   0  'Transparent
      Caption         =   "GERENTE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   44
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbLoja 
      BackStyle       =   0  'Transparent
      Caption         =   "Loja"
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
      Left            =   5640
      TabIndex        =   43
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbVrFixo 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Fixo"
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
      Left            =   5640
      TabIndex        =   42
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbPremio 
      BackStyle       =   0  'Transparent
      Caption         =   "Prêmio"
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
      Left            =   7320
      TabIndex        =   41
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbComis 
      BackStyle       =   0  'Transparent
      Caption         =   "Comissão"
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
      Left            =   7320
      TabIndex        =   40
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbVrMinimo 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Mínimo"
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
      Left            =   5640
      TabIndex        =   39
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   2085
      TabIndex        =   32
      Top             =   1560
      Width           =   600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "®"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   225
      TabIndex        =   31
      Top             =   2265
      Width           =   720
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "@"
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
      Left            =   1440
      TabIndex        =   30
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "ANOTAÇÃO"
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
      Left            =   240
      TabIndex        =   29
      Top             =   5745
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NOME"
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
      Left            =   240
      TabIndex        =   28
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "(D)"
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
      Left            =   2880
      TabIndex        =   27
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "(F)"
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
      Left            =   240
      TabIndex        =   26
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVAÇÃO"
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
      Left            =   225
      TabIndex        =   25
      Top             =   4635
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "V.  PISO"
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
      Left            =   255
      TabIndex        =   24
      Top             =   6675
      Width           =   960
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   255
      TabIndex        =   23
      Top             =   3015
      Width           =   600
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "V.  PISO ®"
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
      Left            =   1545
      TabIndex        =   22
      Top             =   6675
      Width           =   1200
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "QT. P/ MÉDIA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2865
      TabIndex        =   21
      Top             =   6675
      Width           =   1050
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Vcto (F)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      TabIndex        =   20
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Paga salário família"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   7440
      Width           =   1620
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº FILHOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   375
      TabIndex        =   18
      Top             =   7800
      Width           =   960
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   7575
      Left            =   5400
      Top             =   840
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   7575
      Left            =   0
      Top             =   840
      Width           =   5295
   End
End
Attribute VB_Name = "frm_Alt_Funcionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W_LD_FILTRO As Boolean
Dim w_ado_Central As ADODB.Recordset
Dim w_cpf_old As String
Dim w_ativar As Boolean

Private Sub ck_pg_SFam_Click()
    If ck_pg_SFam.Enabled And ck_pg_SFam.value Then
        txt_NFilhos.Enabled = True
    Else
        txt_NFilhos.Enabled = False
    End If
End Sub

Private Sub cmdDescCalcFixo_Click()
    w_CodFunc = frm_Alt_Funcionario.txtFCod
    frm_Alt_Desc_Calc_fixo.lbFunc.Caption = TXT_NOME.text
    frm_Alt_Desc_Calc_fixo.Show 1
End Sub

Sub Form_Load()
On Error Resume Next
w_ativar = False

    'Se estiver Fechada ,  Abre as tabelas
    'If de.rsTAB_FUNC_CRED.State = 0 Then de.TAB_FUNC_CRED
    If de.rsTAB_FUNC_CENTRAL.State = 0 Then de.TAB_FUNC_CENTRAL

On Error GoTo err1
    
    
    'If de.rsTAB_FUNCIONARIO.State = 1 Then de.rsTAB_FUNCIONARIO.Close
    'de.TAB_FUNCIONARIO
    
    If de.rscmdSqlAltFunc.State = 1 Then de.rscmdSqlAltFunc.Close
    de.cmdSqlAltFunc
    Set adoReg.Recordset = de.rscmdSqlAltFunc.Clone
    
    'Funcionarios da Central
    'Set w_ado_Central = de.rsTAB_FUNC_CENTRAL.Clone   'de.cnc.Execute("SELECT COD_LOJ + MID(STR(INT(COD_FUNC)), 2) AS CODIGO, NOME, APELIDO, CARGO, COD_LOJ, COD_FUNC FROM lojb011 ORDER BY NOME, COD_LOJ + MID(STR(INT(COD_FUNC)), 2)").Clone
    'Set ADO_CENTRAL.Recordset = w_ado_Central.Clone
    
    
If acessoTotal() Then
    
    TXT_LOJA.Visible = True
    CK_PREMIO.Visible = True
    TXT_VR_FIXO.Visible = True
    txtFCod.Visible = True
    TXT_COMIS.Visible = True
    TXT_VR_MINIMO.Visible = True
    txt_notas.Visible = True
    Shape2.Visible = True
    lbLoja.Visible = True
    lbPremio.Visible = True
    lbVrFixo.Visible = True
    lbComis.Visible = True
    lbVrMinimo.Visible = True
    lbNotas.Visible = True
    lbGerente.Visible = True
    
End If
    
sair:

    'Set ADOREG.Recordset = de.rsTAB_FUNCIONARIO.Clone  'de.cnc.Execute("select * from tab_funcionario order by f_nome")
    'Set adoReg.Recordset = de.cnc.Execute("SELECT [NUM] & ' ' & [F_COD_L] AS B, tab_funcionario.* FROM tab_funcionario INNER JOIN Lojb010 ON tab_funcionario.F_Cod_L = Lojb010.COD_LOJ;").Clone
    'adoReg.Recordset.Sort = "f_nome"
    
    
    
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

'*** Caption no navegador ***
Private Sub adoReg_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1

If Not IsNumeric(adoReg.Recordset.RecordCount) Then adoReg.Caption = "REGISTRO : " & adoReg.Recordset.AbsolutePosition & " / " & adoReg.Recordset.RecordCount & IIf(W_LD_FILTRO = True, " (FILTRADO)", "")

    Select Case adoReg.Recordset.Fields("F_TIPO")
        Case "V": txt_tipo = "VENDEDOR"
        Case "G": txt_tipo = "GERENTE"
        Case "D": txt_tipo = "GER RODA"
        Case "C": txt_tipo = "CAIXA"
        Case "2": txt_tipo = "2º CAIXA"
        Case "X": txt_tipo = "CX EXTRA"
        Case "R": txt_tipo = "SEGURANÇA"
        Case "S": txt_tipo = "SUPERVISOR"
        Case "O": txt_tipo = "RP"
    End Select

If IsNull(adoReg.Recordset.Fields("F_DT_DEM")) Then
    BarraF.Buttons("editar").Enabled = True
    BarraF.Buttons("ativar").Enabled = False
Else
    If txtCPF.text = "" Then
        BarraF.Buttons("editar").Enabled = True
    Else
        BarraF.Buttons("editar").Enabled = False
    End If
    BarraF.Buttons("ativar").Enabled = True
End If



sair:
    Exit Sub
err1:
    If Not Err.Number = -2147217885 And Not Err.Number = -2147467259 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
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
        Case "ativar": Ativar
        Case "nova":
            If txtCPF.text = "" Then
                MsgBox "Não é possível criar uma Nova Ficha sem CPF! Digite um CPF válido.", vbCritical, "CPF Inválido"
                Exit Sub
            End If
            w_Func_atual = adoReg.Recordset.Fields("F_CODIGO")
            'If ((IsNull(de.cnc.Execute("SELECT F_DT_DEM FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & w_Func_atual).Fields(0))) And (de.cnc.Execute("SELECT COUNT(M_NFICHA) FROM TAB_FICHA_MENS WHERE M_F_COD = " & w_Func_atual).Fields(0)) > 0) Then
            '    MsgBox "Ainda existem fichas abertas para o funcionário " & UCase(ADOREG.Recordset.Fields("F_NOME")) & "!", vbCritical, "Não foi possível criar Nova Ficha"
            'Else
                frm_Cad_Fic_Mensal.Show 1
            'End If
    End Select
End Sub


'*** Rotinas ***
Private Sub Cancelar()
On Error GoTo err1
'If txt_ANOTACAO.Enabled = True Then

    pos = adoReg.Recordset.AbsolutePosition - 1
    adoReg.Recordset.CancelBatch adAffectCurrent
    W_FILTRO = adoReg.Recordset.Filter
    Editar
    adoReg.Refresh
    adoReg.Recordset.Filter = W_FILTRO
    
    'adoReg.Recordset.Sort = "f_nome"
    adoReg.Recordset.Move pos


    
'End If
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub Ativar()
On Error GoTo err1

w_ativar = True

    Editar
    TXT_DT_DEM.Enabled = True
    TXT_DT_DEM = ""
    TXT_DT_DEM.Enabled = False

    TXT_LOGO2.SetFocus

w_ativar = False

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Editar()
On Error GoTo err1
If Not adoReg.Recordset.EOF Then
    w_cpf_old = txtCPF
    BarraF.Buttons("salvar").Enabled = Not BarraF.Buttons("salvar").Enabled
    BarraF.Buttons("cancelar").Enabled = Not BarraF.Buttons("cancelar").Enabled
    BarraF.Buttons("editar").Enabled = Not BarraF.Buttons("editar").Enabled
    
    Grid.Enabled = Not Grid.Enabled
    txt_tipo.Enabled = Not txt_tipo.Enabled
    TXT_NOME.Enabled = Not TXT_NOME.Enabled
    TXT_DT_REG.Enabled = Not TXT_DT_REG.Enabled
    txt_DT_ADM.Enabled = Not txt_DT_ADM.Enabled
    If acessoTotal Then TXT_DT_DEM.Enabled = Not TXT_DT_DEM.Enabled
    TXT_ANOTACAO.Enabled = Not TXT_ANOTACAO.Enabled
    TXT_LOGO.Enabled = Not TXT_LOGO.Enabled
    TXT_LOGO2.Enabled = Not TXT_LOGO2.Enabled
    'TXT_CRED.Enabled = Not TXT_CRED.Enabled
    TXT_FERIAS.Enabled = Not TXT_FERIAS.Enabled
    TXT_OBS.Enabled = Not TXT_OBS.Enabled
    txt_notas.Enabled = Not txt_notas.Enabled
    txt_VPiso.Enabled = Not txt_VPiso.Enabled
    txt_VPiso_R.Enabled = Not txt_VPiso_R.Enabled
    txt_Vcto_ferias.Enabled = Not txt_Vcto_ferias.Enabled
    ck_pg_SFam.Enabled = Not ck_pg_SFam.Enabled
    ck_pg_vt.Enabled = Not ck_pg_vt.Enabled
    dtNasc.Enabled = Not dtNasc.Enabled
    
    If ck_pg_SFam.Enabled And ck_pg_SFam.value Then
        txt_NFilhos.Enabled = True
    Else
        txt_NFilhos.Enabled = False
    End If
    
    
    TXT_LOJA.Enabled = Not TXT_LOJA.Enabled
    CK_PREMIO.Enabled = Not CK_PREMIO.Enabled
    TXT_VR_FIXO.Enabled = Not TXT_VR_FIXO.Enabled
    TXT_COMIS.Enabled = Not TXT_COMIS.Enabled
    TXT_VR_MINIMO.Enabled = Not TXT_VR_MINIMO.Enabled
    
    txt_cod_central.Enabled = Not txt_cod_central.Enabled
    
    txtCPF.Enabled = Not txtCPF.Enabled
    'ElseIf txtCPF = "" Or (UCase(w_usuario) = UCase(NomeMestre) Or UCase(w_usuario) = UCase(NomeMestre2) Or UCase(w_usuario) = UCase(NomeMestre3)) Then
    '    txtCPF.Enabled = True
    'End If
    
    
    'If TXT_LOGO <> "" Then TXT_CENTRAL.Enabled = Not TXT_CENTRAL.Enabled
    

'FILTRO OS DADOS SOMENTE DA LOJA DO REGISTRO
    'TXT_LOGO_Validate False

    If BarraF.Buttons("salvar").Enabled = False Then
        Grid.SetFocus
    Else
        'TXT_CRED.SetFocus
    End If
    
    
    w_ck_vt = ck_pg_vt
    
    
    
Else
    MsgBox "Não existe registro para editar!", vbExclamation
End If
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Excluir()
On Error GoTo err1
  
If Not adoReg.Recordset.EOF Then
    
    If vbYes = MsgBox("DESEJA REALMENTE EXCLUIR O FUNCIONÁRIO (" & TXT_NOME & ")?", vbQuestion + vbYesNo) Then
        adoReg.Recordset.Delete
        adoReg.Recordset.UpdateBatch
        Form_Load

    End If
   
 Else
    MsgBox "Não existe registro para excluir!", vbExclamation
End If
   
sair:
    Exit Sub
err1:
    If Not Err.Number = -2147467259 Then
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    Else
        MsgBox "NÃO É POSSÍVEL EXCLUIR ESTE FUNCIONÁRIO, DEVIDO A FICHAS MENSAIS RELACIONADAS A ELE!", vbCritical
        adoReg.Refresh
    End If
    Resume sair
End Sub

Private Sub Fechar()
On Error GoTo err1
    de.rsTAB_FUNCIONARIO.Requery
    de.rsTAB_FUNCIONARIO.Close
sair:
    Unload Me
    Exit Sub
err1:
'    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub FILTRAR()
Dim w_resp As String
Dim W_CAMPO As String
Dim W_FILTRO As String

On Error GoTo err1
    
    w_resp = InputBox("FILTRAR PELO QUÊ ? ENTRE COM O NÚMERO DA OPÇÃO DESEJADA." & Chr(13) & Chr(13) & "1 - NOME" & Chr(13) & "2 - LOGO" & Chr(13) & "3 - DATA ADMISSÃO" & Chr(13) & "4 - DATA DE REGISTRO" & Chr(13) & "5 - DATA DE DEMISSÃO" & Chr(13) & "6 - ADMITIDOS" & Chr(13) & "7 - CPF" & Chr(13) & "8 - CÓDIGO MILLENNIUM" & Chr(13) & "9 - REMOVER FILTRO *", , "1")
    
    
    If Not w_resp = "" And IsNumeric(w_resp) And w_resp >= 1 And w_resp <= 9 Then
        Select Case w_resp
        'NOME
        Case 1:
            w_resp = "NOME"
            W_CAMPO = "F_NOME"
        'LOGO
        Case 2:
            w_resp = "LOGO"
            W_CAMPO = "B"
        'DT_ADM
        Case 3:
            w_resp = "DT_ADM"
            W_CAMPO = "F_DT_ADM"
        'DT_REG
        Case 4:
            w_resp = "DT_REG"
            W_CAMPO = "F_DT_REG"
        'DT_DEM
        Case 5:
            w_resp = "DT_DEM"
            W_CAMPO = "F_DT_DEM"
        'ADM
        Case 6:
            w_resp = "ADM"
            W_CAMPO = "F_DT_DEM"
        Case 7:
            w_resp = "CPF"
            W_CAMPO = "F_CPF"
        Case 8:
            w_resp = "Código Millennium"
            W_CAMPO = "F_COD_CENTRAL"
        
        Case 9:
            If Not adoReg.Recordset.Filter = 0 Then
                W_LD_FILTRO = False
                adoReg.Recordset.Filter = 0
                adoReg.Refresh
            End If
        End Select
        If Not w_resp = "9" Then
            
            If Mid(w_resp, 1, 2) = "DT" Then
                frm_ESCOLHA_DATA.Show 1
                W_FILTRO = W_CAMPO & " >= #" & frm_ESCOLHA_DATA.TXT_DT_INICIAL & "# AND " & W_CAMPO & " <= #" & frm_ESCOLHA_DATA.TXT_DT_FINAL & "#"
                W_LD_FILTRO = True
                adoReg.Recordset.Filter = W_FILTRO
            ElseIf w_resp = "ADM" Then
                W_FILTRO = "(" & W_CAMPO & ") = NULL"
                W_LD_FILTRO = True
                adoReg.Recordset.Filter = W_FILTRO
            Else
                W_FILTRO = InputBox("ENTRE COM O " & w_resp & " DESEJADO !")
                W_FILTRO = W_CAMPO & " LIKE '%" & W_FILTRO & "%'"
                W_LD_FILTRO = True
                adoReg.Recordset.Filter = W_FILTRO
            End If
        End If
    End If
    
sair:
    Exit Sub
err1:
    If Err.Number <> 13 And Err.Number <> 3265 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
        W_LD_FILTRO = False
        Resume sair

End Sub

Private Sub Salvar()
Dim w_func As String
On Error GoTo err1
'If txt_ANOTACAO.Enabled = True Then
    
     If txtCPF <> w_cpf_old Then
        If Not calculacpf(txtCPF.text) Then
            MsgBox "CPF incorreto! Digite novamente.", vbCritical, "CPF Inválido"
            txtCPF = ""
            txtCPF.SetFocus
            Exit Sub
        End If
    
        w_func = ""
        w_func = de.cnc.Execute("SELECT iif(First(F_NOME),First(F_NOME),'NULL') FROM TAB_FUNCIONARIO WHERE F_CPF = '" & txtCPF & "'").Fields(0)
        'Verifica se já existe o CPF cadastrado
        If (w_func <> "NULL") Then
            MsgBox "O CPF " & txtCPF & " já existe no funcionário " & UCase(w_func) & ". Verifique!", vbCritical, "CPF já existe"
            txtCPF.SetFocus
            Exit Sub
        End If
    End If
    
    
    adoReg.Recordset.UpdateBatch adAffectCurrent
    
    Dim fichaAtual As String
    fichaAtual = de.cnc.Execute("SELECT Max(M_NFICHA) FROM TAB_FICHA_MENS GROUP BY TAB_FICHA_MENS.M_F_COD HAVING (((TAB_FICHA_MENS.M_F_COD)= " & adoReg.Recordset.Fields("F_CODIGO") & "))").Fields(0)
    de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_NOME = '" & TXT_NOME & "', M_LOGO = '" & TXT_LOGO & "', M_PG_VT = '" & ck_pg_vt & "' WHERE (M_F_COD = " & adoReg.Recordset.Fields("F_CODIGO") & " AND M_NFICHA = " & fichaAtual & " )", w_reg
    
    'Pagto Vale Transporte
    'If ck_pg_vt <> w_ck_vt Then
    If (ck_pg_vt = 1 And w_ck_vt = False) Or (ck_pg_vt = 0 And w_ck_vt = True) Then
        If ck_pg_vt Then
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_PG_VT = 1 WHERE (F_Codigo = " & adoReg.Recordset.Fields("F_CODIGO") & " )", w_reg
            de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & fichaAtual & " AND (C_TP_CONTA = 109 OR C_TP_CONTA = 110 OR C_TP_CONTA = 111)")
            de.cnc.Execute ("DELETE FROM TAB_DESC_CALC_FIXO WHERE CF_EMP_COD = " & adoReg.Recordset.Fields("F_CODIGO") & " AND (CF_TP_CONTA = 109 OR CF_TP_CONTA = 110 OR CF_TP_CONTA = 111)")
                
                Dim adoFixos As ADODB.Recordset
                
                Dim ultimoFixo As String
               
                de.cmdIncluirDescCalcFixo Now(), adoReg.Recordset.Fields("F_CODIGO"), "109", "-", "0", "INSS 8% do piso [GERADO AUTOMATICAMENTE]"
                ultimoFixo = de.cnc.Execute("SELECT Max([CF_CODIGO]) FROM TAB_DESC_CALC_FIXO").Fields(0)
                Set adoFixos = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_CODIGO = " & ultimoFixo).Clone
                de.cmdIncluirDescCalc2 Date, fichaAtual, adoFixos.Fields("CF_TP_CONTA"), adoFixos.Fields("CF_TP_OP"), adoFixos.Fields("CF_VALOR"), adoFixos.Fields("CF_DESC"), "0", adoFixos.Fields("CF_CODIGO"), "0", "0", adoFixos.Fields("CF_EMP_COD"), 0

                ultimoFixo = Empty
                Set adoFixos = Nothing
            
                de.cmdIncluirDescCalcFixo Now(), adoReg.Recordset.Fields("F_CODIGO"), "110", "-", "0", "Vale Transporte 6% do piso [GERADO AUTOMATICAMENTE]"
                ultimoFixo = de.cnc.Execute("SELECT Max([CF_CODIGO]) FROM TAB_DESC_CALC_FIXO").Fields(0)
                Set adoFixos = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_CODIGO = " & ultimoFixo).Clone
                de.cmdIncluirDescCalc2 Date, fichaAtual, adoFixos.Fields("CF_TP_CONTA"), adoFixos.Fields("CF_TP_OP"), adoFixos.Fields("CF_VALOR"), adoFixos.Fields("CF_DESC"), "0", adoFixos.Fields("CF_CODIGO"), "0", "0", adoFixos.Fields("CF_EMP_COD"), 0

                ultimoFixo = Empty
                Set adoFixos = Nothing
            
                de.cmdIncluirDescCalcFixo Now(), adoReg.Recordset.Fields("F_CODIGO"), "111", "=", "0", "Pagto. de passes (vale transporte) [GERADO AUTOMATICAMENTE]"
                ultimoFixo = de.cnc.Execute("SELECT Max([CF_CODIGO]) FROM TAB_DESC_CALC_FIXO").Fields(0)
                Set adoFixos = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_CODIGO = " & ultimoFixo).Clone
                de.cmdIncluirDescCalc2 Date, fichaAtual, adoFixos.Fields("CF_TP_CONTA"), adoFixos.Fields("CF_TP_OP"), adoFixos.Fields("CF_VALOR"), adoFixos.Fields("CF_DESC"), "0", adoFixos.Fields("CF_CODIGO"), "0", "0", adoFixos.Fields("CF_EMP_COD"), 0
            
                fichaAtual = Empty
                ultimoFixo = Empty
                Set adoFixos = Nothing
                
        Else
            MsgBox "Para CANCELAR o Pagamento de Vale Transporte é necessário a senha mestre.", vbInformation, "Confirmação de senha"
            frm_Habilitar.Show 1
            w_PSS = frm_Habilitar.txt_Pss
            If w_PSS = w_PassWordLib Then
                de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_PG_VT = 0 WHERE (F_Codigo = " & adoReg.Recordset.Fields("F_CODIGO") & " )", w_reg
                de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & fichaAtual & " AND (C_TP_CONTA = 109 OR C_TP_CONTA = 110 OR C_TP_CONTA = 111)")
                de.cnc.Execute ("DELETE FROM TAB_DESC_CALC_FIXO WHERE CF_EMP_COD  = " & adoReg.Recordset.Fields("F_CODIGO") & " AND (CF_TP_CONTA = 109 OR CF_TP_CONTA = 110 OR CF_TP_CONTA = 111)")
            Else
                ck_pg_vt = 1
            End If
    
        End If
    End If


    
    
    'de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_NOME = '" & TXT_NOME & "' WHERE (M_F_COD = " & adoReg.Recordset.Fields("F_CODIGO") & ")", w_reg
    
    Editar
    
    'Set ADO_CENTRAL.Recordset = w_ado_Central.Clone
    
    
'End If
  
    If w_ativar Then
    
        w_Func_atual = adoReg.Recordset.Fields("F_CODIGO")
        frm_Cad_Fic_Mensal.Show
    
    End If
  
sair:
    Exit Sub
err1:
    'MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
    
End Sub















Private Sub txt_ANOTACAO_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then
        KeyCode = 0
        Sendkeys "{BACKSPACE}"
        Sendkeys "{tab}"
      End If
End Sub

Private Sub TXT_CENTRAL_Change()
    If Not TXT_CENTRAL_COD.BoundText = TXT_CENTRAL.BoundText And BarraF.Buttons("editar").Enabled = False Then TXT_CENTRAL_COD.BoundText = TXT_CENTRAL.BoundText
End Sub



Private Sub TXT_CENTRAL_GotFocus()
' SendKeys "{F4}"
End Sub

Private Sub TXT_CENTRAL_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{Tab}"

End Sub

Private Sub TXT_CRED_Click(Area As Integer)
On Error GoTo err1
  
  'If TXT_CRED.BoundText <> TXT_CRED.BoundText And BarraF.Buttons("editar").Enabled = False Then TXT_CRED_N.Text = TXT_CRED.BoundText
      
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub





Private Sub TXT_CRED_GotFocus()
'    SendKeys "{F4}"
End Sub



Private Sub txt_Cred_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{Tab}"

End Sub



Private Sub txt_DiasTrab_KeyDown(KeyCode As Integer, Shift As Integer)
 KeyEnter KeyCode
 Pause 0.3
 If KeyCode = 13 Then If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1) Then Salvar

End Sub

Private Sub txt_DT_ADM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{Tab}"

End Sub

Private Sub TXT_DT_DEM_GotFocus()
    Sendkeys "{home}+{end}"
End Sub

Private Sub TXT_DT_DEM_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode

End Sub

Private Sub TXT_DT_DEM_Validate(Cancel As Boolean)
    If Not IsDate(TXT_DT_DEM) And TXT_DT_DEM <> "" Then
        MsgBox "Não é permitido digitar outro dado a não ser data!", vbCritical
        TXT_DT_DEM = ""
    End If
End Sub

Private Sub TXT_DT_REG_GotFocus()
    Sendkeys "{home}+{end}"
End Sub

Private Sub TXT_DT_REG_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode

End Sub

Private Sub TXT_DT_REG_Validate(Cancel As Boolean)
    If Not IsDate(TXT_DT_REG) And TXT_DT_REG <> "" Then
        MsgBox "Não é permitido digitar outro dado a não ser data!", vbCritical
        TXT_DT_REG = ""
    End If
End Sub



Private Sub TXT_FERIAS_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then
        KeyCode = 0
        Sendkeys "{BACKSPACE}"
        Sendkeys "{tab}"
      End If
End Sub

Private Sub TXT_LOGO_Click(Area As Integer)
    TXT_LOGO2.BoundText = TXT_LOGO.BoundText
End Sub

Private Sub TXT_LOGO_GotFocus()
' SendKeys "{F4}"
End Sub

Private Sub TXT_LOGO_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{Tab}"
End Sub



Private Sub ck_pg_SFam_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then Sendkeys "{Tab}"
End Sub

Private Sub TXT_LOGO2_Click(Area As Integer)
    TXT_LOGO.BoundText = TXT_LOGO2.BoundText
End Sub

Private Sub txt_NFilhos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{Tab}"
End Sub

Private Sub txt_Nome_GotFocus()
    Sendkeys "{home}+{end}"
End Sub

'--------- Ao Pressionar uma Tecla -----------

Private Sub ck_pg_SFam_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub txt_NFilhos_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub txt_FERIAS_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub TXT_NOME_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub txt_Nome_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
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
Private Sub txt_ANOTACAO_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_DT_ADM_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_DT_REG_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_DT_DEM_KeyUp(KeyCode As Integer, Shift As Integer)
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
    End Select
End If
End Sub



Private Sub TXT_LOGO_Validate(Cancel As Boolean)
On Error GoTo err1

   'Set ADO_CENTRAL.Recordset = de.cnc.Execute("SELECT COD_LOJ + MID(STR(INT(COD_FUNC)), 2) AS CODIGO, NOME, APELIDO, CARGO, COD_LOJ, COD_FUNC FROM lojb011  WHERE (COD_LOJ = '" & TXT_LOGO & "') ORDER BY NOME, COD_LOJ + MID(STR(INT(COD_FUNC)), 2)").Clone

    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub txt_tipo_Click()
    Select Case txt_tipo
        Case "VENDEDOR": TXT_TIPO2 = "V"
        Case "GERENTE": TXT_TIPO2 = "G"
        Case "GER RODA": TXT_TIPO2 = "D"
        Case "CAIXA": TXT_TIPO2 = "C"
        Case "2º CAIXA": TXT_TIPO2 = "2"
        Case "CX EXTRA": TXT_TIPO2 = "X"
        Case "SEGURANÇA": TXT_TIPO2 = "R"
        Case "SUPERVISOR": TXT_TIPO2 = "S"
        Case "RP": TXT_TIPO2 = "O"
    End Select
End Sub

Private Sub TXT_TIPO_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then Sendkeys "{tab}"
End Sub


Private Sub txt_Vcto_ferias_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txt_VPiso_KeyDown(KeyCode As Integer, Shift As Integer)
 KeyEnter KeyCode
End Sub

Private Sub txt_VPiso_R_KeyDown(KeyCode As Integer, Shift As Integer)
 KeyEnter KeyCode
End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub txtCPF_KeyPress(KeyAscii As Integer)
  'se teclar enter envia um TAB
  If KeyAscii = 13 Then
     Sendkeys "{TAB}"
     KeyAscii = 0
  End If
End Sub

Private Sub txtCPF_LostFocus()
    
    If Len(txtCPF.text) > 0 Then
      Select Case Len(txtCPF.text)
       Case Is = 11
         If Not calculacpf(txtCPF.text) Then
            MsgBox "CPF incorreto! Digite novamente.", vbCritical, "CPF Inválido"
            txtCPF = ""
            txtCPF.SetFocus
         End If
       End Select
    End If
End Sub

Private Sub dtNasc_GotFocus()
    Sendkeys "{home}+{end}"
End Sub

Private Sub dtNasc_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode

End Sub

Private Sub dtNasc_Validate(Cancel As Boolean)
    If Not IsDate(dtNasc) And dtNasc <> "" Then
        MsgBox "Não é permitido digitar outro dado a não ser data!", vbCritical
        dtNasc = ""
    End If
End Sub

