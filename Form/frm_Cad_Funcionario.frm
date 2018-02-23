VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_Cad_Funcionario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Funcionário"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "frm_Cad_Funcionario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox TXT_TIPO 
      Height          =   315
      ItemData        =   "frm_Cad_Funcionario.frx":1CFA
      Left            =   840
      List            =   "frm_Cad_Funcionario.frx":1D19
      TabIndex        =   7
      Text            =   "VENDEDOR"
      Top             =   3120
      Width           =   1875
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
      ItemData        =   "frm_Cad_Funcionario.frx":1D70
      Left            =   120
      List            =   "frm_Cad_Funcionario.frx":1D8F
      TabIndex        =   36
      Text            =   "V"
      Top             =   3120
      Width           =   585
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
      Height          =   285
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3120
      Width           =   2055
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
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   5
      Top             =   2445
      Width           =   1215
   End
   Begin VB.ComboBox txt_Vcto_ferias 
      DataField       =   "F_VCTO_FERIAS"
      DataSource      =   "adoReg"
      Height          =   315
      ItemData        =   "frm_Cad_Funcionario.frx":1DAE
      Left            =   2895
      List            =   "frm_Cad_Funcionario.frx":1DD6
      TabIndex        =   6
      Top             =   2445
      Width           =   810
   End
   Begin VB.TextBox dtNasc 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
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
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox ck_pg_vt 
      Caption         =   "Check1"
      DataField       =   "F_PG_VT"
      DataSource      =   "ADOREG"
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   8280
      Width           =   195
   End
   Begin VB.CommandButton cmdVerificaCPF 
      Caption         =   "?"
      Height          =   255
      Left            =   1680
      TabIndex        =   29
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox ck_pg_SFam 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   8280
      Width           =   195
   End
   Begin VB.TextBox txt_NFilhos 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
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
      Left            =   240
      MaxLength       =   10
      TabIndex        =   15
      Top             =   8760
      Width           =   1095
   End
   Begin VB.TextBox txt_VPiso_R 
      Alignment       =   1  'Right Justify
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   13
      Text            =   "0"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox txt_Vend 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "@@.@@"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   300
      Left            =   3840
      TabIndex        =   24
      Top             =   8760
      Width           =   990
   End
   Begin VB.TextBox txt_VPiso 
      Alignment       =   1  'Right Justify
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      MaxLength       =   10
      TabIndex        =   12
      Text            =   "0"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox TXT_OBS 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   5025
      Width           =   4815
   End
   Begin VB.TextBox TXT_FERIAS 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   3840
      Width           =   4815
   End
   Begin VB.TextBox TXT_NOME 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   4815
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4440
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
            Picture         =   "frm_Cad_Funcionario.frx":1E01
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Cad_Funcionario.frx":211B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Cad_Funcionario.frx":2435
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Cad_Funcionario.frx":274F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Cad_Funcionario.frx":2A69
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Cad_Funcionario.frx":2D83
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt_ANOTACAO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   6225
      Width           =   4815
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   1482
      ButtonWidth     =   1667
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Fechar"
            Key             =   "fechar"
            Object.ToolTipText     =   "Fechar (Alt+F)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "salvar"
            Object.ToolTipText     =   "Salvar Alteração (Alt+S)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar Alteração (Alt+C)"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox txtCPF 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   14
      Format          =   "###.###.###-##"
      Mask            =   "###.###.###-##"
      PromptChar      =   "_"
   End
   Begin MSDataListLib.DataCombo TXT_LOGO 
      Bindings        =   "frm_Cad_Funcionario.frx":309D
      DataField       =   "F_COD_L"
      DataSource      =   "ADOREG"
      Height          =   360
      Left            =   2040
      TabIndex        =   32
      Top             =   1800
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   635
      _Version        =   393216
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
   Begin MSDataListLib.DataCombo TXT_LOGO2 
      Bindings        =   "frm_Cad_Funcionario.frx":30AE
      DataField       =   "F_COD_L"
      DataSource      =   "ADOREG"
      Height          =   360
      Left            =   2880
      TabIndex        =   2
      Top             =   1800
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   635
      _Version        =   393216
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
      Left            =   120
      TabIndex        =   4
      Top             =   2445
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   635
      _Version        =   393216
      Format          =   73007105
      CurrentDate     =   38432
   End
   Begin VB.Label Label2 
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
      Left            =   120
      TabIndex        =   38
      Top             =   2880
      Width           =   600
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
      Left            =   2880
      TabIndex        =   37
      Top             =   2880
      Width           =   2055
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
      Left            =   1425
      TabIndex        =   35
      Top             =   2160
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
      Left            =   120
      TabIndex        =   34
      Top             =   2175
      Width           =   1095
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
      Left            =   2880
      TabIndex        =   33
      Top             =   2175
      Width           =   855
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
      Left            =   3720
      TabIndex        =   31
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
      Left            =   2520
      TabIndex        =   30
      Top             =   8280
      Width           =   1740
   End
   Begin VB.Label Label7 
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
      Left            =   120
      TabIndex        =   28
      Top             =   1560
      Width           =   375
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
      Left            =   120
      TabIndex        =   27
      Top             =   8280
      Width           =   1740
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
      Left            =   255
      TabIndex        =   26
      Top             =   8535
      Width           =   960
   End
   Begin VB.Label Label13 
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
      Left            =   1455
      TabIndex        =   25
      Top             =   7440
      Width           =   1080
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
      Left            =   135
      TabIndex        =   23
      Top             =   7440
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
      Left            =   120
      TabIndex        =   22
      Top             =   4785
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
      Left            =   120
      TabIndex        =   21
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label4 
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
      Left            =   120
      TabIndex        =   20
      Top             =   960
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
      Left            =   2070
      TabIndex        =   19
      Top             =   1560
      Width           =   375
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
      Left            =   150
      TabIndex        =   17
      Top             =   5985
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   8415
      Left            =   0
      Top             =   840
      Width           =   5055
   End
End
Attribute VB_Name = "frm_Cad_Funcionario"
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
    TXT_NOME = ""
    TXT_ANOTACAO = ""
    txt_DT_ADM = Date
    TXT_DT_REG = ""
    TXT_LOGO = ""
    TXT_CENTRAL = ""
    TXT_OBS = ""
    TXT_FERIAS = ""
    TXT_CRED = ""
    txt_tipo = "VENDEDOR"
    txt_VPiso_R = Format(0, "R$ 0.00")
    txt_VPiso = Format(0, "R$ 0.00")
    txt_NFilhos = 0
    ck_pg_SFam = 0
    ck_pg_vt = 0
    TXT_LOJA = ""
    CK_PREMIO = 0
    TXT_VR_FIXO = Format(0, "R$ 0.00")
    TXT_COMIS = Format(0, "0.0")
    TXT_VR_MINIMO = Format(0, "R$ 0.00")
    dtNasc = ""
    
    TXT_NOME.SetFocus
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Fechar()
On Error GoTo err1
    If de.rsTAB_FUNCIONARIO.State = 1 Then de.rsTAB_FUNCIONARIO.Close
    de.TAB_FUNCIONARIO
    Unload Me
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub Salvar()
Dim w_func As String

On Error GoTo err1
    
    If txtCPF = "" Then
        MsgBox "CPF em branco! Digite novamente.", vbCritical, "CPF em branco"
        Exit Sub
    End If
    
    
    If Not calculacpf(txtCPF.Text) Then
            MsgBox "CPF incorreto! Digite novamente.", vbCritical, "CPF Inválido"
            txtCPF = ""
            txtCPF.SetFocus
            Exit Sub
    End If
    
    If TXT_LOGO = "" Then
            MsgBox "(B) em branco! Selecione algum e tente novamente.", vbCritical, "B inválido"
            TXT_LOGO = ""
            TXT_LOGO.SetFocus
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
    
    
    Dim w_tipo

    Select Case txt_tipo
        Case "VENDEDOR": w_tipo = "V"
        Case "GERENTE": w_tipo = "G"
        Case "GER RODA": w_tipo = "D"
        Case "CAIXA": w_tipo = "C"
        Case "2º CAIXA": w_tipo = "2"
        Case "CX EXTRA": w_tipo = "X"
        Case "SEGURANÇA": w_tipo = "R"
        Case "SUPERVISOR": w_tipo = "S"
        Case "RP": w_tipo = "O"
    End Select
   
    de.cmdIncluirFuncionario TXT_NOME, TXT_ANOTACAO, TXT_LOGO, txt_DT_ADM, TXT_FERIAS, TXT_OBS, txt_VPiso, w_tipo, 0, txt_VPiso_R, txt_Vcto_ferias, ck_pg_SFam, txt_NFilhos, Format(TXT_LOJA, "000"), 0, 0, 0, 0, txtCPF, ck_pg_vt, dtNasc
    v_cod = de.cnc.Execute("SELECT MAX(F_CODIGO)AS COD FROM TAB_FUNCIONARIO").Fields(0)
    
    If TXT_DT_REG <> "" And IsDate(TXT_DT_REG) Then
        de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_DT_REG = '" & TXT_DT_REG & "'  WHERE (F_Codigo = " & v_cod & ")"
    End If
    
    'Pagto Vale Transporte
    'Dim fichaAtual As String
    'fichaAtual = de.cnc.Execute("SELECT Max(M_NFICHA) FROM TAB_FICHA_MENS GROUP BY TAB_FICHA_MENS.M_F_COD HAVING (((TAB_FICHA_MENS.M_F_COD)= " & v_cod & "))").Fields(0)
    'If ck_pg_vt Then
    '        Dim adoFixos As ADODB.Recordset
            
    '        Dim ultimoFixo As String
    '
    '        de.cmdIncluirDescCalcFixo Now(), v_cod, "109", "-", "0", "INSS 8% do piso [GERADO AUTOMATICAMENTE]"
    '        ultimoFixo = de.cnc.Execute("SELECT Max([CF_CODIGO]) FROM TAB_DESC_CALC_FIXO").Fields(0)
    '        Set adoFixos = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_CODIGO = " & ultimoFixo).Clone
    '        de.cmdIncluirDescCalc2 Date, fichaAtual, adoFixos.Fields("CF_TP_CONTA"), adoFixos.Fields("CF_TP_OP"), adoFixos.Fields("CF_VALOR"), adoFixos.Fields("CF_DESC"), "0", adoFixos.Fields("CF_CODIGO"), "0", "0", adoFixos.Fields("CF_EMP_COD"), 0

    '        ultimoFixo = Empty
    '        Set adoFixos = Nothing
    '
    '        de.cmdIncluirDescCalcFixo Now(), v_cod, "110", "-", "0", "Vale Transporte 6% do piso [GERADO AUTOMATICAMENTE]"
    '        ultimoFixo = de.cnc.Execute("SELECT Max([CF_CODIGO]) FROM TAB_DESC_CALC_FIXO").Fields(0)
    '        Set adoFixos = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_CODIGO = " & ultimoFixo).Clone
    '        de.cmdIncluirDescCalc2 Date, fichaAtual, adoFixos.Fields("CF_TP_CONTA"), adoFixos.Fields("CF_TP_OP"), adoFixos.Fields("CF_VALOR"), adoFixos.Fields("CF_DESC"), "0", adoFixos.Fields("CF_CODIGO"), "0", "0", adoFixos.Fields("CF_EMP_COD"), 0
    '
    '        fichaAtual = Empty
    '        ultimoFixo = Empty
    '        Set adoFixos = Nothing
    '
    '        'Lancamentos
    'Else
    '    MsgBox "Para CANCELAR o Pagamento de Vale Transporte é necessário a senha mestre.", vbInformation, "Confirmação de senha"
    '    frm_Habilitar.Show 1
    '    w_PSS = frm_Habilitar.txt_Pss
    '    If w_PSS = w_PassWordLib Then
    '        de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & fichaAtual & " AND (C_TP_CONTA = 109 OR C_TP_CONTA = 110)")
    '        de.cnc.Execute ("DELETE FROM TAB_DESC_CALC_FIXO WHERE CF_EMP_COD  = " & v_cod & " AND (CF_TP_CONTA = 109 OR CF_TP_CONTA = 110)")
    '    Else
    '        ck_pg_vt = 1
    '    End If

    'End If

    'w_ck_vt = ck_pg_vt
    
    
    
    MsgBox "Registro salvo com sucesso!", vbInformation
    
    
    w_Func_atual = v_cod
    frm_Cad_Fic_Mensal.txt_DT_ADM = txt_DT_ADM
    frm_Cad_Fic_Mensal.TXT_DT_REG = TXT_DT_REG
    frm_Cad_Fic_Mensal.Show 1
    
    
    Cancelar
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
    
End Sub





Private Sub ck_pg_SFam_Click()
    If ck_pg_SFam.Enabled And ck_pg_SFam.value Then
        txt_NFilhos.Enabled = True
    Else
        txt_NFilhos.Enabled = False
    End If
End Sub

Private Sub cmdVerificaCPF_Click()
   If txtCPF = "" Then
        MsgBox "CPF em branco! Digite novamente.", vbCritical, "CPF em branco"
        txtCPF.SetFocus
        Exit Sub
    End If
    
    
    If Not calculacpf(txtCPF.Text) Then
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
    
    MsgBox "O CPF " & txtCPF & " é válido e está disponível.", vbInformation, "CPF ok"
    
        
End Sub

Private Sub Form_Activate()
 Cancelar
End Sub



Private Sub Form_Load()



txt_Vcto_ferias.ListIndex = 0

End Sub

Private Sub txt_ANOTACAO_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{BACKSPACE}"
        SendKeys "{tab}"
      End If
End Sub



Private Sub TXT_CENTRAL_GotFocus()
    SendKeys "{F4}"
End Sub

Private Sub TXT_CENTRAL_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub TXT_CENTRAL_Validate(Cancel As Boolean)
    txt_Vend = TXT_CENTRAL.BoundText
End Sub

Private Sub TXT_CRED_GotFocus()
     SendKeys "{F4}"
End Sub

Private Sub txt_Cred_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{tab}"
      End If
End Sub






Private Sub txt_DiasTrab_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then
        If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo) Then Salvar
        SendKeys "{tab}"
     End If
End Sub

Private Sub txt_DT_ADM_Change()
    If TXT_DT_REG = "" Then txt_Vcto_ferias = Format(txt_DT_ADM, "MM")
End Sub

Private Sub txt_DT_ADM_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{tab}"
      End If
End Sub

Private Sub TXT_DT_REG_Change()
    If IsDate(TXT_DT_REG) Then txt_Vcto_ferias = Format(TXT_DT_REG, "MM")
End Sub

Private Sub TXT_DT_REG_GotFocus()
    SendKeys "{home}+{end}"
End Sub



Private Sub TXT_DT_REG_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub TXT_FERIAS_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{BACKSPACE}"
        SendKeys "{tab}"
      End If
End Sub


Private Sub TXT_LOGO_Click(Area As Integer)
    TXT_LOGO2.BoundText = TXT_LOGO.BoundText
End Sub

Private Sub TXT_LOGO_GotFocus()
' SendKeys "{F4}"
End Sub

Private Sub TXT_LOGO_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{Tab}"
End Sub








Private Sub ck_pg_SFam_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub TXT_LOGO2_Click(Area As Integer)
    TXT_LOGO.BoundText = TXT_LOGO2.BoundText
End Sub

Private Sub txt_NFilhos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub
Private Sub txt_Nome_GotFocus()
    SendKeys "{home}+{end}"
End Sub

'--------- Ao Pressionar uma Tecla -----------
Private Sub ck_pg_SFam_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub txt_NFilhos_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub TXT_LOGO_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub TXT_NOME_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub TXT_NOME_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_Nome_KeyUp(KeyCode As Integer, Shift As Integer)
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

Private Sub TXT_OBS_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{BACKSPACE}"
        SendKeys "{tab}"
      End If
End Sub

Private Sub txt_OBS_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub txt_FERIAS_KeyUp(KeyCode As Integer, Shift As Integer)
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



Private Sub TXT_LOGO_Validate(Cancel As Boolean)
On Error GoTo err1

    'Set ADO_CENTRAL.Recordset = de.cnc.Execute("SELECT COD_LOJ + Format(MID(STR(INT(COD_FUNC)), 2),'00') AS CODIGO, NOME FROM lojb011 WHERE COD_LOJ = '" & TXT_LOGO & "' ORDER BY NOME, COD_LOJ + MID(STR(INT(COD_FUNC)), 2)").Clone
    
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
      If KeyCode = 13 Then SendKeys "{tab}"
End Sub


Private Sub txt_VPiso_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub
Private Sub txt_VPiso_R_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub txtCPF_KeyPress(KeyAscii As Integer)
  'se teclar enter envia um TAB
  If KeyAscii = 13 Then
     SendKeys "{TAB}"
     KeyAscii = 0
  End If
End Sub

Private Sub txtCPF_LostFocus()
    
    If Len(txtCPF.Text) > 0 Then
      Select Case Len(txtCPF.Text)
       Case Is = 11
         If Not calculacpf(txtCPF.Text) Then
            MsgBox "CPF incorreto! Digite novamente.", vbCritical, "CPF Inválido"
            txtCPF = ""
            txtCPF.SetFocus
         End If
       End Select
    End If
    
End Sub

Private Sub dtNasc_GotFocus()
    SendKeys "{home}+{end}"
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

