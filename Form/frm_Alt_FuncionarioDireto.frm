VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_Alt_FuncionarioDireto 
   Caption         =   "Alteração de Funcionário"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   9720
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
      TabIndex        =   68
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox ck_pg_vt 
      Caption         =   "Check1"
      DataField       =   "F_PG_VT"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   65
      Top             =   7440
      Width           =   195
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
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2055
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
      ItemData        =   "frm_Alt_FuncionarioDireto.frx":0000
      Left            =   240
      List            =   "frm_Alt_FuncionarioDireto.frx":001F
      TabIndex        =   45
      Text            =   "V"
      Top             =   3240
      Width           =   585
   End
   Begin VB.ComboBox TXT_TIPO 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frm_Alt_FuncionarioDireto.frx":003E
      Left            =   960
      List            =   "frm_Alt_FuncionarioDireto.frx":005D
      TabIndex        =   44
      Text            =   "VENDEDOR"
      Top             =   3240
      Width           =   1875
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
      TabIndex        =   12
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CheckBox ck_pg_SFam 
      Caption         =   "Check1"
      DataField       =   "F_PG_SAL_FAM"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   7440
      Width           =   195
   End
   Begin VB.ComboBox txt_Vcto_ferias 
      DataField       =   "F_VCTO_FERIAS"
      DataSource      =   "adoReg"
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frm_Alt_FuncionarioDireto.frx":00B4
      Left            =   4335
      List            =   "frm_Alt_FuncionarioDireto.frx":00DC
      TabIndex        =   10
      Top             =   2550
      Width           =   810
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
      Height          =   300
      Left            =   8400
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   960
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
      TabIndex        =   7
      Top             =   6915
      Width           =   1095
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
      TabIndex        =   6
      Top             =   4890
      Width           =   4890
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
      TabIndex        =   5
      Top             =   3840
      Width           =   4890
   End
   Begin VB.TextBox TXT_NOME 
      DataField       =   "F_NOME"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Width           =   4935
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
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   5970
      Width           =   4875
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
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   1
      Top             =   2550
      Width           =   1215
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
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   880
      Width           =   855
   End
   Begin MSAdodcLib.Adodc adoReg 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   8520
      Width           =   9720
      _ExtentX        =   17145
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5760
      Top             =   120
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
            Picture         =   "frm_Alt_FuncionarioDireto.frx":0107
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_FuncionarioDireto.frx":0421
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_FuncionarioDireto.frx":073B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_FuncionarioDireto.frx":0A55
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_FuncionarioDireto.frx":0D6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_FuncionarioDireto.frx":1089
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   1482
      ButtonWidth     =   1667
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
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
      EndProperty
   End
   Begin MSDataListLib.DataCombo TXT_LOGO 
      Bindings        =   "frm_Alt_FuncionarioDireto.frx":13A3
      DataField       =   "F_COD_L"
      DataSource      =   "ADOREG"
      Height          =   360
      Left            =   2040
      TabIndex        =   14
      Top             =   1920
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
      Height          =   360
      Left            =   240
      TabIndex        =   15
      Top             =   2550
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   219545601
      CurrentDate     =   38432
   End
   Begin TabDlg.SSTab tabFunc 
      Height          =   5415
      Left            =   5400
      TabIndex        =   31
      Top             =   840
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BackColor       =   -2147483648
      TabCaption(0)   =   "GERENTE"
      TabPicture(0)   =   "frm_Alt_FuncionarioDireto.frx":13B4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lbVrMinimo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbComis"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbPremio"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbVrFixo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbLoja"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbNotas"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CK_PREMIO"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TXT_VR_MINIMO"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TXT_COMIS"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TXT_VR_FIXO"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TXT_LOJA"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txt_notas"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "CAIXA"
      TabPicture(1)   =   "frm_Alt_FuncionarioDireto.frx":13D0
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblCxLoja"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblCxFixo"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblCxComis1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblCxMinimo"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblCxComis2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblCxComis3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblCxComisDez"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtCxLoja"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtCxFixo"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtCxComis1"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtCxMinimo"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtCxComis2"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtCxComis3"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtCxComisDez"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      Begin VB.TextBox txtCxComisDez 
         DataField       =   "F_PERC_FIXO_DEZ"
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
         Left            =   240
         TabIndex        =   57
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txtCxComis3 
         DataField       =   "F_COMIS3"
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
         Left            =   1920
         TabIndex        =   56
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtCxComis2 
         DataField       =   "F_COMIS2"
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
         Left            =   240
         TabIndex        =   55
         Top             =   3720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtCxMinimo 
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
         Left            =   240
         TabIndex        =   53
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtCxComis1 
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
         Left            =   240
         TabIndex        =   54
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtCxFixo 
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
         Left            =   240
         TabIndex        =   52
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtCxLoja 
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
         Left            =   240
         TabIndex        =   51
         Top             =   840
         Width           =   855
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
         Height          =   2295
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   3000
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
         Left            =   -74760
         TabIndex        =   34
         Top             =   840
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
         Left            =   -74760
         TabIndex        =   35
         Top             =   1560
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
         Left            =   -73080
         TabIndex        =   36
         Top             =   1560
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
         Left            =   -74760
         TabIndex        =   37
         Top             =   2280
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox CK_PREMIO 
         Caption         =   "Check1"
         DataField       =   "F_PREMIO"
         DataSource      =   "adoReg"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -73080
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblCxComisDez 
         BackStyle       =   0  'Transparent
         Caption         =   "% Comis Dez."
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
         TabIndex        =   64
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label lblCxComis3 
         BackStyle       =   0  'Transparent
         Caption         =   "% Comis 3"
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
         Left            =   1920
         TabIndex        =   63
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblCxComis2 
         BackStyle       =   0  'Transparent
         Caption         =   "% Comis 2"
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
         TabIndex        =   62
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblCxMinimo 
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
         Left            =   240
         TabIndex        =   61
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblCxComis1 
         BackStyle       =   0  'Transparent
         Caption         =   "% Comis 1"
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
         TabIndex        =   60
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblCxFixo 
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
         Left            =   240
         TabIndex        =   59
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblCxLoja 
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
         Left            =   240
         TabIndex        =   58
         Top             =   600
         Width           =   1335
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
         Left            =   -74760
         TabIndex        =   30
         Top             =   2640
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   -74760
         TabIndex        =   43
         Top             =   600
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
         Left            =   -74760
         TabIndex        =   42
         Top             =   1320
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
         Left            =   -73080
         TabIndex        =   41
         Top             =   600
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
         Left            =   -73080
         TabIndex        =   40
         Top             =   1320
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
         Left            =   -74760
         TabIndex        =   39
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin MSMask.MaskEdBox txtCPF 
      DataField       =   "F_CPF"
      DataSource      =   "adoReg"
      Height          =   360
      Left            =   240
      TabIndex        =   48
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
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
      Bindings        =   "frm_Alt_FuncionarioDireto.frx":13EC
      DataField       =   "F_COD_L"
      DataSource      =   "ADOREG"
      Height          =   360
      Left            =   2880
      TabIndex        =   50
      Top             =   1920
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
      TabIndex        =   67
      Top             =   1680
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
      TabIndex        =   66
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
      TabIndex        =   49
      Top             =   1680
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
      Left            =   3000
      TabIndex        =   47
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "COD:"
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
      TabIndex        =   32
      Top             =   885
      Width           =   495
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
      Height          =   375
      Left            =   375
      TabIndex        =   29
      Top             =   7800
      Width           =   960
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
      TabIndex        =   28
      Top             =   7440
      Width           =   1740
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
      Left            =   4320
      TabIndex        =   27
      Top             =   2280
      Width           =   855
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
      TabIndex        =   26
      Top             =   6675
      Width           =   1200
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
      Left            =   240
      TabIndex        =   25
      Top             =   3000
      Width           =   600
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
      TabIndex        =   23
      Top             =   4635
      Width           =   1815
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
      TabIndex        =   22
      Top             =   3600
      Width           =   1815
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
      Left            =   3000
      TabIndex        =   21
      Top             =   2280
      Width           =   1065
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
      TabIndex        =   20
      Top             =   1080
      Width           =   1335
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
      TabIndex        =   19
      Top             =   5745
      Width           =   1335
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
      Left            =   240
      TabIndex        =   18
      Top             =   2280
      Width           =   1095
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
      Left            =   1545
      TabIndex        =   17
      Top             =   2265
      Width           =   720
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
      TabIndex        =   16
      Top             =   1680
      Width           =   600
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   7575
      Left            =   120
      Top             =   840
      Width           =   5175
   End
End
Attribute VB_Name = "frm_Alt_FuncionarioDireto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W_LD_FILTRO As Boolean
Dim w_ado_Central As ADODB.Recordset
Dim w_cpf_old As String
Dim w_ck_vt As Boolean


Private Sub ck_pg_SFam_Click()
    If ck_pg_SFam.Enabled And ck_pg_SFam.value Then
        txt_NFilhos.Enabled = True
    Else
        txt_NFilhos.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    
w_cpf_old = ""

    
On Error GoTo err1
    
    
    'If de.rsTAB_FUNCIONARIO.State = 1 Then de.rsTAB_FUNCIONARIO.Close
    'de.TAB_FUNCIONARIO
  
    
    'Funcionarios da Central
    'Set w_ado_Central = de.rsTAB_FUNC_CENTRAL.Clone   'de.cnc.Execute("SELECT COD_LOJ + MID(STR(INT(COD_FUNC)), 2) AS CODIGO, NOME, APELIDO, CARGO, COD_LOJ, COD_FUNC FROM lojb011 ORDER BY NOME, COD_LOJ + MID(STR(INT(COD_FUNC)), 2)").Clone
    'Set ADO_CENTRAL.Recordset = w_ado_Central.Clone
    
    If de.rscmdBase.State = 1 Then de.rscmdBase.Close
    
    de.rscmdBase.Open "Select * from TAB_FUNCIONARIO where F_CODIGO = " & frm_Alt_Fic_Mensal_VIS.txt_F_COD
    Set adoReg.Recordset = de.rscmdBase.Clone
    de.rscmdBase.Close
    
    If TXT_TIPO2.text = "C" Then
        tabFunc.Tab = 1
    ElseIf TXT_TIPO2.text = "G" Or TXT_TIPO2.text = "D" Or TXT_TIPO2.text = "X" Then
        tabFunc.Tab = 0
    Else
        tabFunc.Visible = False
    End If
    
    
If acessoTotal() Then
    
    TXT_LOJA.Visible = True
    CK_PREMIO.Visible = True
    TXT_VR_FIXO.Visible = True
    'txtFCod.Visible = True
    TXT_COMIS.Visible = True
    TXT_VR_MINIMO.Visible = True
    txt_notas.Visible = True
    'Shape2.Visible = True
    lbLoja.Visible = True
    lbPremio.Visible = True
    lbVrFixo.Visible = True
    lbComis.Visible = True
    lbVrMinimo.Visible = True
    lbNotas.Visible = True
    'lbGerente.Visible = True
    
Else
    If TXT_TIPO2 = "G" Or TXT_TIPO2 = "D" Or TXT_TIPO2 = "X" Then
        tabFunc.Visible = False
    End If
End If
    
sair:
    
    'Set adoReg.Recordset = de.rsTAB_FUNCIONARIO.Clone  'de.cnc.Execute("select * from tab_funcionario order by f_nome")
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
    End Select
End Sub


'*** Rotinas ***
Private Sub Cancelar()
On Error GoTo err1
'If txt_ANOTACAO.Enabled = True Then

    pos = adoReg.Recordset.AbsolutePosition - 1
    adoReg.Recordset.CancelBatch adAffectCurrent
    'W_FILTRO = adoReg.Recordset.Filter
    Editar
    'adoReg.Refresh
    'adoReg.Recordset.Filter = W_FILTRO
    
    'adoReg.Recordset.Sort = "f_nome"
    'adoReg.Recordset.Move pos


    
'End If
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub Editar()
On Error GoTo err1
If Not adoReg.Recordset.EOF Then

    w_ck_vt = ck_pg_vt
    w_cpf_old = txtCPF
    BarraF.Buttons("salvar").Enabled = Not BarraF.Buttons("salvar").Enabled
    BarraF.Buttons("cancelar").Enabled = Not BarraF.Buttons("cancelar").Enabled
    BarraF.Buttons("editar").Enabled = Not BarraF.Buttons("editar").Enabled
    
    
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
    
    txtCxLoja.Enabled = Not txtCxLoja.Enabled
    txtCxFixo.Enabled = Not txtCxFixo.Enabled
    txtCxMinimo.Enabled = Not txtCxMinimo.Enabled
    txtCxComis1.Enabled = Not txtCxComis1.Enabled
    txtCxComis2.Enabled = Not txtCxComis2.Enabled
    txtCxComis3.Enabled = Not txtCxComis3.Enabled
    txtCxComisDez.Enabled = Not txtCxComisDez.Enabled
    
    txt_cod_central.Enabled = Not txt_cod_central.Enabled
    
    If txtCPF.Enabled Then
        txtCPF.Enabled = False
    ElseIf txtCPF = "" Or (UCase(w_usuario) = UCase(NomeMestre) Or UCase(w_usuario) = UCase(NomeMestre2) Or UCase(w_usuario) = UCase(NomeMestre3)) Then
        txtCPF.Enabled = True
    End If
    
    'If TXT_LOGO <> "" Then TXT_CENTRAL.Enabled = Not TXT_CENTRAL.Enabled
    

'FILTRO OS DADOS SOMENTE DA LOJA DO REGISTRO
    'TXT_LOGO_Validate False

    If BarraF.Buttons("salvar").Enabled = False Then
        'Grid.SetFocus
    Else
        'TXT_CRED.SetFocus
    End If
    
    
    
    
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

    frm_Alt_Fic_Mensal_VIS.Timer1 = True

    'de.rsTAB_FUNCIONARIO.Requery
    'de.rsTAB_FUNCIONARIO.Close
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
    
    w_resp = InputBox("FILTRAR PELO QUÊ ? ENTRE COM O NÚMERO DA OPÇÃO DESEJADA." & Chr(13) & Chr(13) & "1 - NOME" & Chr(13) & "2 - LOGO" & Chr(13) & "3 - DATA ADMISSÃO" & Chr(13) & "4 - DATA DE REGISTRO" & Chr(13) & "5 - DATA DE DEMISSÃO" & Chr(13) & "6 - ADMITIDOS" & Chr(13) & "7 - REMOVER FILTRO *", , "1")
    
    
    If Not w_resp = "" And IsNumeric(w_resp) And w_resp >= 1 And w_resp <= 7 Then
        Select Case w_resp
        'NOME
        Case 1:
            w_resp = "NOME"
            W_CAMPO = "F_NOME"
        'LOGO
        Case 2:
            w_resp = "LOGO"
            W_CAMPO = "F_Cod_L"
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
            If Not adoReg.Recordset.Filter = 0 Then
                W_LD_FILTRO = False
                adoReg.Recordset.Filter = 0
                adoReg.Refresh
            End If
        End Select
        If Not w_resp = "7" Then
            
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
            MsgBox "O CPF " & txtCPF & " já está cadastrado. Verifique!", vbCritical, "CPF já existe"
            txtCPF.SetFocus
            Exit Sub
        End If
    End If
    
    
    adoReg.Recordset.UpdateBatch adAffectCurrent
    
    Dim fichaAtual As String
    fichaAtual = de.cnc.Execute("SELECT Max(M_NFICHA) FROM TAB_FICHA_MENS GROUP BY TAB_FICHA_MENS.M_F_COD HAVING (((TAB_FICHA_MENS.M_F_COD)= " & adoReg.Recordset.Fields("F_CODIGO") & "))").Fields(0)
    de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_NOME = '" & TXT_NOME & "', M_LOGO = '" & TXT_LOGO & "', M_PG_VT = '" & ck_pg_vt & "' , M_TIPO = '" & TXT_TIPO2 & "'WHERE (M_F_COD = " & adoReg.Recordset.Fields("F_CODIGO") & " AND M_NFICHA = " & fichaAtual & " )", w_reg
    
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
                
                'Lancamentos
        Else
            MsgBox "Para CANCELAR o Pagamento de Vale Transporte é necessário a senha mestre.", vbInformation, "Confirmação de senha"
            frm_Habilitar.Show 1
            w_PSS = frm_Habilitar.txt_Pss
            If w_PSS = w_PassWordLib Then
                de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_PG_VT = 0 WHERE (F_Codigo = " & adoReg.Recordset.Fields("F_CODIGO") & " )", w_reg
                de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & fichaAtual & " AND (C_TP_CONTA = 109 OR C_TP_CONTA = 110 OR C_TP_CONTA = 111)")
                de.cnc.Execute ("DELETE FROM TAB_DESC_CALC_FIXO WHERE CF_EMP_COD  = " & adoReg.Recordset.Fields("F_CODIGO") & " AND (CF_TP_CONTA = 109 OR CF_TP_CONTA = 110 R CF_TP_CONTA = 111)")
            Else
                ck_pg_vt = 1
            End If
    
        End If
    End If

    
    Editar
    
    'Set ADO_CENTRAL.Recordset = w_ado_Central.Clone
    
    
'End If
  
sair:
    Exit Sub
err1:
    'MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
    
End Sub

Private Sub tab_DblClick()

End Sub

Private Sub Form_Unload(Cancel As Integer)
            Fechar
End Sub

Private Sub txt_ANOTACAO_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then
        KeyCode = 0
        Sendkeys "{BACKSPACE}"
        Sendkeys "{tab}"
      End If
End Sub

Private Sub TXT_CENTRAL_Change()
    'If Not TXT_CENTRAL_COD.BoundText = TXT_CENTRAL.BoundText And BarraF.Buttons("editar").Enabled = False Then TXT_CENTRAL_COD.BoundText = TXT_CENTRAL.BoundText
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









Private Sub TXT_CRED_N_Click(Area As Integer)

End Sub

Private Sub TXT_COMIS_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then Sendkeys "{tab}"
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

Private Sub TXT_LOJA_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then Sendkeys "{tab}"
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
    'Case 84: ' "T"
            'FILTRAR
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

Private Sub TXT_VR_FIXO_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub TXT_VR_MINIMO_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then Sendkeys "{tab}"
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

Private Sub txtCxComis1_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txtCxComis2_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txtCxComis3_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txtCxComisDez_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txtCxFixo_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txtCxLoja_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txtCxMinimo_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then Sendkeys "{tab}"
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

