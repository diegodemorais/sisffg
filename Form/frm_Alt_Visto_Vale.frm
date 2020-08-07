VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "msCOMCTL.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form frm_Alt_Visto_Vale 
   AutoRedraw      =   -1  'True
   Caption         =   "Vistar Contas"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   14445
   Icon            =   "frm_Alt_Visto_Vale.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   14445
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmCargos 
      BackColor       =   &H80000003&
      Caption         =   " Cargo "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   10320
      TabIndex        =   51
      Top             =   960
      Width           =   3015
      Begin VB.CheckBox ckTipo 
         BackColor       =   &H80000002&
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
         Height          =   255
         Left            =   1920
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
      End
      Begin VB.ListBox txt_tipo 
         Height          =   1815
         ItemData        =   "frm_Alt_Visto_Vale.frx":030A
         Left            =   120
         List            =   "frm_Alt_Visto_Vale.frx":0329
         MultiSelect     =   1  'Simple
         TabIndex        =   52
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame frmGeral 
      BackColor       =   &H80000003&
      Caption         =   " Geral "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5400
      TabIndex        =   44
      Top             =   960
      Width           =   4695
      Begin VB.CheckBox ckZerados 
         BackColor       =   &H80000003&
         Caption         =   "Mostrar Zerados"
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
         Height          =   255
         Left            =   2400
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox ckFixos 
         BackColor       =   &H80000003&
         Caption         =   "Programados dentro do mês informado"
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
         Height          =   495
         Left            =   2400
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox TXT_ANO 
         Alignment       =   1  'Right Justify
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
         Left            =   1200
         TabIndex        =   48
         Top             =   600
         Width           =   810
      End
      Begin VB.ComboBox TXT_MES 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_Alt_Visto_Vale.frx":0380
         Left            =   240
         List            =   "frm_Alt_Visto_Vale.frx":03A8
         TabIndex        =   46
         Text            =   "TXT_MES"
         Top             =   600
         Width           =   780
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "ANO"
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
         Left            =   1200
         TabIndex        =   47
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "MÊS"
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
         TabIndex        =   45
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame frmContas 
      BackColor       =   &H80000003&
      Caption         =   " Conta "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5400
      TabIndex        =   40
      Top             =   2160
      Width           =   4695
      Begin VB.CheckBox ckConta 
         BackColor       =   &H80000002&
         Caption         =   "Todas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   600
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox TXT_CONTA_COD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin MSDataListLib.DataCombo TXT_CONTA 
         Bindings        =   "frm_Alt_Visto_Vale.frx":03D3
         Height          =   360
         Left            =   1080
         TabIndex        =   42
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "TP_DESC"
         BoundColumn     =   "TP_COD"
         Text            =   "%"
         Object.DataMember      =   "SQL_TP_CONTA"
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
   End
   Begin VB.Frame frmFuncionario 
      BackColor       =   &H80000003&
      Caption         =   " Funcionário "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   35
      Top             =   1800
      Width           =   4935
      Begin VB.PictureBox btnLoadingFuncionario 
         Height          =   975
         Left            =   1920
         Picture         =   "frm_Alt_Visto_Vale.frx":03E4
         ScaleHeight     =   915
         ScaleWidth      =   915
         TabIndex        =   54
         Top             =   200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox ck_Nome 
         BackColor       =   &H80000002&
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
         Height          =   255
         Left            =   3840
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   960
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CommandButton cmdNome 
         Caption         =   ">"
         Enabled         =   0   'False
         Height          =   310
         Left            =   4560
         TabIndex        =   38
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtNome 
         Enabled         =   0   'False
         Height          =   310
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   4335
      End
      Begin MSDataListLib.DataCombo dbNome 
         Bindings        =   "frm_Alt_Visto_Vale.frx":04EE
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "F_NOME"
         BoundColumn     =   "F_Codigo"
         Text            =   "%"
         Object.DataMember      =   ""
      End
      Begin MSAdodcLib.Adodc ADO_CENTRAL 
         Height          =   330
         Left            =   1680
         Top             =   960
         Visible         =   0   'False
         Width           =   1260
         _ExtentX        =   2223
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
         Caption         =   "CENTRAL"
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
   End
   Begin MSAdodcLib.Adodc ADO_FUNC 
      Height          =   375
      Left            =   3360
      Top             =   840
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin VB.Frame frmLojas 
      BackColor       =   &H80000003&
      Caption         =   " (B) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   31
      Top             =   960
      Width           =   3015
      Begin VB.CheckBox ckTodas 
         BackColor       =   &H80000002&
         Caption         =   "Todas"
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
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin MSDataListLib.DataCombo TXT_LOGO 
         Bindings        =   "frm_Alt_Visto_Vale.frx":0505
         DataField       =   "F_COD_L"
         DataSource      =   "ADOREG"
         Height          =   360
         Left            =   240
         TabIndex        =   32
         Top             =   240
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
         Bindings        =   "frm_Alt_Visto_Vale.frx":0516
         DataField       =   "F_COD_L"
         DataSource      =   "ADOREG"
         Height          =   360
         Left            =   1080
         TabIndex        =   33
         Top             =   240
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
   End
   Begin VB.TextBox txt_NVistT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
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
      Height          =   285
      Left            =   12645
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   7440
      Width           =   1620
   End
   Begin VB.TextBox txt_VistT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
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
      Height          =   285
      Left            =   12645
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   7080
      Width           =   1620
   End
   Begin VB.TextBox TXT_TOTALT 
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
      Height          =   285
      Left            =   12645
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   7800
      Width           =   1620
   End
   Begin VB.TextBox TXT_TOTALIgual 
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
      Height          =   285
      Left            =   9045
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   7800
      Width           =   1620
   End
   Begin VB.TextBox txt_VistIgual 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
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
      Height          =   285
      Left            =   9045
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   7080
      Width           =   1620
   End
   Begin VB.TextBox txt_NVistIgual 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
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
      Height          =   285
      Left            =   9045
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   7440
      Width           =   1620
   End
   Begin VB.TextBox TXT_TOTALMenos 
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
      Height          =   285
      Left            =   5445
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   7800
      Width           =   1620
   End
   Begin VB.TextBox txt_VistMenos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
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
      Height          =   285
      Left            =   5445
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   7080
      Width           =   1620
   End
   Begin VB.TextBox txt_NVistMenos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
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
      Height          =   285
      Left            =   5445
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   7440
      Width           =   1620
   End
   Begin MSAdodcLib.Adodc adoConta 
      Height          =   330
      Left            =   120
      Top             =   6600
      Visible         =   0   'False
      Width           =   2850
      _ExtentX        =   5027
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
   Begin VB.TextBox txt_NVistMais 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
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
      Height          =   285
      Left            =   1845
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   7440
      Width           =   1620
   End
   Begin VB.TextBox txt_VistMais 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
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
      Height          =   285
      Left            =   1845
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   7080
      Width           =   1620
   End
   Begin VB.CommandButton cmdPesq 
      BackColor       =   &H80000002&
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13440
      Picture         =   "frm_Alt_Visto_Vale.frx":0527
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox TXT_TOTALMais 
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
      Height          =   285
      Left            =   1845
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   7800
      Width           =   1620
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
            Picture         =   "frm_Alt_Visto_Vale.frx":33A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Visto_Vale.frx":36BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Visto_Vale.frx":39D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Visto_Vale.frx":3CEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Visto_Vale.frx":4009
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Visto_Vale.frx":4323
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Visto_Vale.frx":463D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   14445
      _ExtentX        =   25479
      _ExtentY        =   1482
      ButtonWidth     =   1667
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
            Caption         =   "&Imprimir"
            Key             =   "imprimir"
            Object.ToolTipText     =   "Imprime a Ficha (Alt+I)"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid grid_Conta 
      Bindings        =   "frm_Alt_Visto_Vale.frx":4A8F
      Height          =   3570
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3360
      Width           =   14250
      _ExtentX        =   25135
      _ExtentY        =   6297
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      Enabled         =   0   'False
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
      ColumnCount     =   17
      BeginProperty Column00 
         DataField       =   "LOGO"
         Caption         =   "LOGO"
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
      BeginProperty Column02 
         DataField       =   "C_TP_CONTA"
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "DATA"
         Caption         =   "DT LCTO"
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
      BeginProperty Column05 
         DataField       =   "DESCR"
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
      BeginProperty Column06 
         DataField       =   "VALOR"
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
      BeginProperty Column07 
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
      BeginProperty Column08 
         DataField       =   "VISTO"
         Caption         =   "VISTO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "Sim"
            FalseValue      =   "Não"
            NullValue       =   "Não"
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "C_NCRED"
         Caption         =   "C_NCRED"
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
      BeginProperty Column10 
         DataField       =   "codigo"
         Caption         =   "codigo"
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
         DataField       =   "Ficha"
         Caption         =   "Ficha"
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
      BeginProperty Column13 
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
      BeginProperty Column14 
         DataField       =   "FUNC"
         Caption         =   "FUNC"
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
      BeginProperty Column15 
         DataField       =   "C_DATA_INTERNA"
         Caption         =   "C_DATA_INTERNA"
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
      BeginProperty Column16 
         DataField       =   "TP_COD"
         Caption         =   "TP_COD"
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
         Locked          =   -1  'True
         BeginProperty Column00 
            Alignment       =   2
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
         EndProperty
         BeginProperty Column09 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column10 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column11 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column12 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column13 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column14 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column15 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column16 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7380
      TabIndex        =   30
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   7290
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10950
      TabIndex        =   29
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   10890
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3810
      TabIndex        =   28
      Top             =   7200
      Width           =   255
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   3690
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   165
      TabIndex        =   27
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   90
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Não Vistado"
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
      Index           =   10
      Left            =   11520
      TabIndex        =   26
      Top             =   7485
      Width           =   1050
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vistado"
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
      Index           =   9
      Left            =   11520
      TabIndex        =   25
      Top             =   7125
      Width           =   645
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL "
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
      Index           =   8
      Left            =   11520
      TabIndex        =   24
      Top             =   7845
      Width           =   675
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL "
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
      Left            =   7920
      TabIndex        =   20
      Top             =   7845
      Width           =   675
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vistado"
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
      Left            =   7920
      TabIndex        =   19
      Top             =   7125
      Width           =   645
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Não Vistado"
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
      Left            =   7920
      TabIndex        =   18
      Top             =   7485
      Width           =   1050
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL "
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
      Left            =   4320
      TabIndex        =   14
      Top             =   7845
      Width           =   675
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vistado"
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
      Left            =   4320
      TabIndex        =   13
      Top             =   7125
      Width           =   645
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Não Vistado"
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
      Left            =   4320
      TabIndex        =   12
      Top             =   7485
      Width           =   1050
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Não Vistado"
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
      Left            =   720
      TabIndex        =   8
      Top             =   7485
      Width           =   1050
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vistado"
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
      Left            =   720
      TabIndex        =   6
      Top             =   7125
      Width           =   645
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      Height          =   2415
      Left            =   120
      Top             =   840
      Width           =   14250
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL "
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
      Left            =   720
      TabIndex        =   4
      Top             =   7845
      Width           =   675
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   480
      Top             =   6960
      Width           =   3135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   4080
      Top             =   6960
      Width           =   3135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   7680
      Top             =   6960
      Width           =   3135
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   11280
      Top             =   6960
      Width           =   3135
   End
   Begin VB.Menu mnu 
      Caption         =   "Menu"
      Begin VB.Menu mnuVist 
         Caption         =   "Vistar"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuVolt 
         Caption         =   "Remover"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnutodos 
         Caption         =   "Vistar Todos"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnutodosR 
         Caption         =   "Remover Todos"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "frm_Alt_Visto_Vale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W_LD_FILTRO As Boolean
Dim W_POS As Long
Dim w_Move As Boolean
Dim W_Pv  As Boolean ' SE É A PRIMEIRA VEZ QUE ENTREA NA TELA
Dim W_FILTRO As String
Dim wStrSql As String  'SQL  da Consulta
Dim w_load As Boolean

'Flag para la tecla BackSpace
Private bKeyBack As Boolean


Public Function AutoCompletar_TextBox()
    If (Len(txtNome.text) = 0) Then
        Exit Function
    End If
        'btnLoadingFuncionario.Visible = True
        ADO_FUNC.Recordset.MoveFirst
        Do Until (ADO_FUNC.Recordset.EOF)
            If InStr(1, ADO_FUNC.Recordset.Fields("F_NOME"), txtNome.text, vbTextCompare) = 1 Then
                dbNome.text = ADO_FUNC.Recordset.Fields("F_NOME")
                'btnLoadingFuncionario.Visible = False
                Exit Function
            End If
            ADO_FUNC.Recordset.MoveNext
        Loop
    'btnLoadingFuncionario.Visible = False
    dbNome = "%"
End Function


Private Sub cmdNome_Click()
    If txtNome.Enabled And dbNome.Enabled Then
        AutoCompletar_TextBox
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub



Private Sub Picture1_Click()

End Sub

Private Sub txtNome_Change()
    Dim selSt As Long
    selSt = txtNome.SelStart
    txtNome.text = UCase(txtNome.text)
    txtNome.SelStart = selSt
End Sub

Private Sub txtNome_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyBack, vbKeyDelete
            Select Case Len(txtNome.text)
                Case Is <> 0
                    bKeyBack = True
            End Select
        Case 13
            Call AutoCompletar_TextBox
            txtNome.SetFocus
    End Select
End Sub



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
            If ADO_TOTAL.Fields("C_TP_Op") = "+" Then
                TXT_MAIS = CDbl(TXT_MAIS) + ADO_TOTAL.Fields("C_VALOR")
            ElseIf ADO_TOTAL.Fields("C_TP_Op") = "-" Then
                TXT_MENOS = CDbl(TXT_MENOS) + ADO_TOTAL.Fields("C_VALOR")
            End If
            ADO_TOTAL.MoveNext
        Loop
        
        TXT_TOTAL = CDbl(TXT_MAIS) - CDbl(TXT_MENOS)
    End If
    
    TXT_MAIS = Format(TXT_MAIS, "R$ #0.00")
    TXT_MENOS = Format(TXT_MENOS, "R$ #0.00")
    TXT_TOTAL = Format(TXT_TOTAL, "R$ #0.00")
    
    'muda cor do total
    If TXT_TOTAL < 0 Then
        TXT_TOTAL.ForeColor = vbRed
    Else
        TXT_TOTAL.ForeColor = vbBlue
    End If

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub ck_Nome_Click()
    If ck_Nome.value = 1 Then
        dbNome = "%"
        txtNome = ""
        txtNome.Enabled = False
        dbNome.Enabled = False
        cmdNome.Enabled = False
        
    Else
        ckTodas.value = 1
        txtNome.Enabled = True
        dbNome.Enabled = True
        cmdNome.Enabled = True
        On Error Resume Next
        txtNome.SetFocus
        Sendkeys "{f4}"
        
    End If
End Sub

Private Sub ck_Nome_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then Sendkeys "{tab}"

End Sub

Private Sub ckConta_Click()

    If ckConta.value = 1 Then
        TXT_CONTA = "%"
        TXT_CONTA.Enabled = False
        TXT_CONTA_cod.Enabled = False
       
    Else
        TXT_CONTA = ""
        TXT_CONTA.Enabled = True
        TXT_CONTA_cod.Enabled = True
        On Error Resume Next
        TXT_CONTA.SetFocus
        Sendkeys "{f4}"
      
    End If
    
End Sub

Private Sub ckConta_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then Sendkeys "{tab}"

End Sub

Private Sub ckFixos_Click()
    cmdPesq_Click
End Sub

Private Sub ckTipo_Click()
    If ckTipo.value = 1 Then
        For I = 0 To txt_tipo.ListCount - 1
            txt_tipo.Selected(I) = True
        Next
    Else
        For I = 0 To txt_tipo.ListCount - 1
            txt_tipo.Selected(I) = False
        Next
    End If
End Sub

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


'** Barra de Ferramenta ***
Private Sub BarraF_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
        Case "fechar": Fechar
        Case "imprimir": Imprimir
    End Select
End Sub


'*** Rotinas ***
Sub Imprimir()
Dim w_Conta As String
Dim w_tp_conta

On Error GoTo err1
If Len(TXT_CONTA.text) > 0 Then w_Conta = Mid(TXT_CONTA.text, 1, Len(TXT_CONTA.text) - (Len(TXT_CONTA_cod.text) + 4))
    
    If TXT_CONTA_cod = "" Then
        w_tp_conta = "%"
    Else
        w_tp_conta = TXT_CONTA_cod
    End If
    
      w_FirstTipo = True

    For J = 0 To txt_tipo.ListCount - 1
        If txt_tipo.Selected(J) = True Then
            Select Case txt_tipo.list(J)
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
        
            If w_FirstTipo Then
                w_Tipos = "" & w_tipo & ""
            Else
                w_Tipos = w_Tipos & "," & w_tipo & ""
            End If
            w_FirstTipo = False
        End If
    Next

    'Verifica se é programado novo (começando a partir do mês)
    If ckFixos Then
        If ck_Nome.value = 0 Then
            If ckZerados = 1 Then
                If de.rscmdSqlVistarFixos2Zerados_Grouping.State = 1 Then de.rscmdSqlVistarFixos2Zerados_Grouping.Close
                de.cmdSqlVistarFixos2Zerados_Grouping dbNome, w_tp_conta, IIf(w_Conta = "", "%", w_Conta), TXT_MES, TXT_ANO, IIf(TXT_LOGO = "", "%", TXT_LOGO), w_Tipos
                rptRelVistoFixoZerados.Show 1
            Else
                If de.rscmdSqlVistarFixos2_Grouping.State = 1 Then de.rscmdSqlVistarFixos2_Grouping.Close
                de.cmdSqlVistarFixos2_Grouping dbNome, w_tp_conta, IIf(w_Conta = "", "%", w_Conta), TXT_MES, TXT_ANO, IIf(TXT_LOGO = "", "%", TXT_LOGO), w_Tipos
                rptRelVistoFixo.Show 1
            End If
        Else
            If ckZerados = 1 Then
                If de.rscmdSqlVistarFixosZerados_Grouping.State = 1 Then de.rscmdSqlVistarFixosZerados_Grouping.Close
                de.cmdSqlVistarFixosZerados_Grouping w_tp_conta, IIf(w_Conta = "", "%", w_Conta), TXT_MES, TXT_ANO, IIf(TXT_LOGO = "", "%", TXT_LOGO), w_Tipos
                rptRelVistoTFixoZerados.Show 1
            Else
                If de.rscmdSqlVistarFixos_Grouping.State = 1 Then de.rscmdSqlVistarFixos_Grouping.Close
                de.cmdSqlVistarFixos_Grouping w_tp_conta, IIf(w_Conta = "", "%", w_Conta), TXT_MES, TXT_ANO, IIf(TXT_LOGO = "", "%", TXT_LOGO), w_Tipos
                rptRelVistoTFixo.Show 1
            End If
        End If
        
    Else

        If ck_Nome.value = 0 Then
            If ckZerados = 1 Then
                If de.rscmdSqlVistar2Zerados_Grouping.State = 1 Then de.rscmdSqlVistar2Zerados_Grouping.Close
                de.cmdSqlVistar2Zerados_Grouping dbNome, w_tp_conta, IIf(w_Conta = "", "%", w_Conta), TXT_MES, TXT_ANO, IIf(TXT_LOGO = "", "%", TXT_LOGO), w_Tipos
                rptRelVistoZerados.Show 1
            Else
                If de.rscmdSqlVistar2_Grouping.State = 1 Then de.rscmdSqlVistar2_Grouping.Close
                de.cmdSqlVistar2_Grouping dbNome, w_tp_conta, IIf(w_Conta = "", "%", w_Conta), TXT_MES, TXT_ANO, IIf(TXT_LOGO = "", "%", TXT_LOGO), w_Tipos
                rptRelVisto.Show 1
            End If
        Else
            If ckZerados = 1 Then
                If de.rscmdSqlVistarZerados_Grouping.State = 1 Then de.rscmdSqlVistarZerados_Grouping.Close
                de.cmdSqlVistarZerados_Grouping w_tp_conta, IIf(w_Conta = "", "%", w_Conta), TXT_MES, TXT_ANO, IIf(TXT_LOGO = "", "%", TXT_LOGO), w_Tipos
                rptRelVistoTZerados.Show 1
            Else
                If de.rscmdSqlVistar_Grouping.State = 1 Then de.rscmdSqlVistar_Grouping.Close
                de.cmdSqlVistar_Grouping w_tp_conta, IIf(w_Conta = "", "%", w_Conta), TXT_MES, TXT_ANO, IIf(TXT_LOGO = "", "%", TXT_LOGO), w_Tipos
                rptRelVistoT.Show 1
            End If
        End If
    End If
          
sair:
    Exit Sub
err1:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume sair
    
End Sub


Private Sub Fechar()
On Error GoTo err1
    'If de.rsTAB_FICHA_MENS.State = 1 Then de.rsTAB_FICHA_MENS.Requery
    Unload Me
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub ckTodas_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then Sendkeys "{tab}"

End Sub

Private Sub cmdCalcIgual_Click()
Dim w_FirtTipo As Boolean
Dim w_Tipos
Dim w_tipo
Dim nenhumTipo As Boolean

On erro GoTo err1

    If txt_tipo.Enabled = True Then
        nenhumTipo = True
        For J = 0 To txt_tipo.ListCount - 1
            If txt_tipo.Selected(J) = True Then
               nenhumTipo = False
            End If
        Next
      
        If nenhumTipo Then
                MsgBox "Selecione ao menos um TIPO!", vbCritical
                Exit Sub
        End If
    End If

    'tipos
    w_FirstTipo = True

    For J = 0 To txt_tipo.ListCount - 1
        If txt_tipo.Selected(J) = True Then
            Select Case txt_tipo.list(J)
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
        
            If w_FirstTipo Then
                w_Tipos = "'" & w_tipo & "'"
            Else
                w_Tipos = w_Tipos & ",'" & w_tipo & "'"
            End If
            w_FirstTipo = False
        End If
    Next
  
    txt_VistIgual.Enabled = True
  
    txt_NVistIgual.Enabled = True
  
    TXT_TOTALIgual.Enabled = True
  
If ((TXT_LOGO = "" And ckTodas.value = 1) Or (TXT_LOGO <> "" And ckTodas.value = 0)) And ((dbNome = "" And ck_Nome.value = 1) Or (dbNome <> "" And ck_Nome.value = 0)) And ((TXT_CONTA = "" And ckConta.value = 1) Or (TXT_CONTA <> "" And ckConta.value = 0)) Then
    If ck_Nome.value = 0 Then
    'um nome
        If dbNome = "" Then
   
            TXT_TOTALIgual = 0
            TXT_TOTALIgual = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP = '='").Fields("VALOR"), "R$ 0.00")
            txt_VistIgual = 0
            txt_VistIgual = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP = '=' and TAB_DESC_CALC.c_Visto = -1").Fields("Valor"), "R$ 0.00")
            txt_NVistIgual = 0
            txt_NVistIgual = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP = '=' and TAB_DESC_CALC.c_Visto = 0").Fields("Valor"), "R$ 0.00")
         
        Else
      
            TXT_TOTALIgual = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR  FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') AND (TAB_FICHA_MENS.M_NOME LIKE '" & dbNome & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP = '='").Fields("Valor"), "R$ 0.00")
            txt_VistIgual = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR  FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') AND (TAB_FICHA_MENS.M_NOME LIKE '" & dbNome & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP = '=' and TAB_DESC_CALC.c_Visto = -1 ").Fields("Valor"), "R$ 0.00")
            txt_NVistIgual = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR  FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') AND (TAB_FICHA_MENS.M_NOME LIKE '" & dbNome & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP = '=' and TAB_DESC_CALC.c_Visto = 0 ").Fields("Valor"), "R$ 0.00")
            
        End If
        
    Else
    'Todos os nomes
    
        TXT_TOTALIgual = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP = '=' AND ((TAB_FUNCIONARIO.F_TIPO) IN (" & w_Tipos & "))").Fields("VALOR"), "R$ 0.00")
        txt_VistIgual = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP = '=' and TAB_DESC_CALC.c_Visto = -1 AND ((TAB_FUNCIONARIO.F_TIPO) IN (" & w_Tipos & "))").Fields("Valor"), "R$ 0.00")
        txt_NVistIgual = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP = '=' and TAB_DESC_CALC.c_Visto = 0 AND ((TAB_FUNCIONARIO.F_TIPO) IN (" & w_Tipos & "))").Fields("Valor"), "R$ 0.00")
      
    End If

    If txt_VistIgual = "" Then txt_VistIgual = "R$ 0,00"
    If txt_NVistIgual = "" Then txt_NVistIgual = "R$ 0,00"
    If TXT_TOTALIgual = "" Then TXT_TOTALIgual = "R$ 0,00"
    
End If


sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub cmdCalcMais_Click()

    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub cmdCalcMenos_Click()
Dim w_FirtTipo As Boolean
Dim w_Tipos
Dim w_tipo
Dim nenhumTipo As Boolean

On erro GoTo err1

    If txt_tipo.Enabled = True Then
        nenhumTipo = True
        For J = 0 To txt_tipo.ListCount - 1
            If txt_tipo.Selected(J) = True Then
               nenhumTipo = False
            End If
        Next
         
        If nenhumTipo Then
                MsgBox "Selecione ao menos um TIPO!", vbCritical
                Exit Sub
        End If
    End If

    'tipos
    w_FirstTipo = True

    For J = 0 To txt_tipo.ListCount - 1
        If txt_tipo.Selected(J) = True Then
            Select Case txt_tipo.list(J)
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
        
            If w_FirstTipo Then
                w_Tipos = "'" & w_tipo & "'"
            Else
                w_Tipos = w_Tipos & ",'" & w_tipo & "'"
            End If
            w_FirstTipo = False
        End If
    Next


    txt_VistMenos.Enabled = True
 
    txt_NVistMenos.Enabled = True
 
    TXT_TOTALMenos.Enabled = True

If ((TXT_LOGO = "" And ckTodas.value = 1) Or (TXT_LOGO <> "" And ckTodas.value = 0)) And ((dbNome = "" And ck_Nome.value = 1) Or (dbNome <> "" And ck_Nome.value = 0)) And ((TXT_CONTA = "" And ckConta.value = 1) Or (TXT_CONTA <> "" And ckConta.value = 0)) Then
    If ck_Nome.value = 0 Then
    'um nome
        If dbNome = "" Then
   
            TXT_TOTALMenos = 0
            TXT_TOTALMenos = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP = '-'").Fields("VALOR"), "R$ 0.00")
            txt_VistMenos = 0
            txt_VistMenos = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP = '-' and TAB_DESC_CALC.c_Visto = -1").Fields("Valor"), "R$ 0.00")
            txt_NVistMenos = 0
            txt_NVistMenos = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP = '-' and TAB_DESC_CALC.c_Visto = 0").Fields("Valor"), "R$ 0.00")
            
        Else
     
            TXT_TOTALMenos = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR  FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') AND (TAB_FICHA_MENS.M_NOME LIKE '" & dbNome & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP = '-'").Fields("Valor"), "R$ 0.00")
            txt_VistMenos = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR  FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') AND (TAB_FICHA_MENS.M_NOME LIKE '" & dbNome & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP = '-' and TAB_DESC_CALC.c_Visto = -1 ").Fields("Valor"), "R$ 0.00")
            txt_NVistMenos = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR  FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') AND (TAB_FICHA_MENS.M_NOME LIKE '" & dbNome & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP = '-' and TAB_DESC_CALC.c_Visto = 0 ").Fields("Valor"), "R$ 0.00")
   
        End If
        
    
    Else
    'Todos os nomes
 
        TXT_TOTALMenos = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP = '-' AND ((TAB_FUNCIONARIO.F_TIPO) IN (" & w_Tipos & "))").Fields("VALOR"), "R$ 0.00")
        txt_VistMenos = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP = '-' and TAB_DESC_CALC.c_Visto = -1 AND ((TAB_FUNCIONARIO.F_TIPO) IN (" & w_Tipos & "))").Fields("Valor"), "R$ 0.00")
        txt_NVistMenos = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP = '-' and TAB_DESC_CALC.c_Visto = 0 AND ((TAB_FUNCIONARIO.F_TIPO) IN (" & w_Tipos & "))").Fields("Valor"), "R$ 0.00")
        
        
    End If

    If txt_VistMenos = "" Then txt_VistMenos = "R$ 0,00"
  
    If txt_NVistMenos = "" Then txt_NVistMenos = "R$ 0,00"
    If TXT_TOTALMenos = "" Then TXT_TOTALMenos = "R$ 0,00"
  
  End If
  
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub cmdCalcT_Click()
Dim w_FirtTipo As Boolean
Dim w_Tipos
Dim w_tipo
Dim nenhumTipo As Boolean

On erro GoTo err1

    If txt_tipo.Enabled = True Then
        nenhumTipo = True
        For J = 0 To txt_tipo.ListCount - 1
            If txt_tipo.Selected(J) = True Then
               nenhumTipo = False
            End If
        Next
        
        If nenhumTipo Then
                MsgBox "Selecione ao menos um TIPO!", vbCritical
                Exit Sub
        End If
    End If

    'tipos
    w_FirstTipo = True

    For J = 0 To txt_tipo.ListCount - 1
        If txt_tipo.Selected(J) = True Then
            Select Case txt_tipo.list(J)
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
        
            If w_FirstTipo Then
                w_Tipos = "'" & w_tipo & "'"
            Else
                w_Tipos = w_Tipos & ",'" & w_tipo & "'"
            End If
            w_FirstTipo = False
        End If
    Next
  
    txt_VistT.Enabled = True
  
    txt_NVistT.Enabled = True
  
    TXT_TOTALT.Enabled = True
    
If ((TXT_LOGO = "" And ckTodas.value = 1) Or (TXT_LOGO <> "" And ckTodas.value = 0)) And ((dbNome = "" And ck_Nome.value = 1) Or (dbNome <> "" And ck_Nome.value = 0)) And ((TXT_CONTA = "" And ckConta.value = 1) Or (TXT_CONTA <> "" And ckConta.value = 0)) Then
    If ck_Nome.value = 0 Then
    'um nome
        If dbNome = "" Then
 
            TXT_TOTALT = 0
            TXT_TOTALT = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP <> '='").Fields("VALOR"), "R$ 0.00")
            txt_VistT = 0
            txt_VistT = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP <> '=' and TAB_DESC_CALC.c_Visto = -1").Fields("Valor"), "R$ 0.00")
            txt_NVistT = 0
            txt_NVistT = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP <> '=' and TAB_DESC_CALC.c_Visto = 0").Fields("Valor"), "R$ 0.00")
            
            
        Else
  
            TXT_TOTALT = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR  FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') AND (TAB_FICHA_MENS.M_NOME LIKE '" & dbNome & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP <> '='").Fields("Valor"), "R$ 0.00")
            txt_VistT = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR  FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') AND (TAB_FICHA_MENS.M_NOME LIKE '" & dbNome & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP <> '=' and TAB_DESC_CALC.c_Visto = -1 ").Fields("Valor"), "R$ 0.00")
            txt_NVistT = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR  FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') AND (TAB_FICHA_MENS.M_NOME LIKE '" & dbNome & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP <> '=' and TAB_DESC_CALC.c_Visto = 0 ").Fields("Valor"), "R$ 0.00")
            
        End If
        
    
    Else
    'Todos os nomes
 
        TXT_TOTALT = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP <> '=' AND ((TAB_FUNCIONARIO.F_TIPO) IN (" & w_Tipos & "))").Fields("VALOR"), "R$ 0.00")
        txt_VistT = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP <> '=' and TAB_DESC_CALC.c_Visto = -1 AND ((TAB_FUNCIONARIO.F_TIPO) IN (" & w_Tipos & "))").Fields("Valor"), "R$ 0.00")
        txt_NVistT = Format(de.cnc.Execute("SELECT sum(TAB_DESC_CALC.C_VALOR) AS VALOR FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_FUNCIONARIO, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " and TAB_DESC_CALC.C_TP_OP <> '=' and TAB_DESC_CALC.c_Visto = 0 AND ((TAB_FUNCIONARIO.F_TIPO) IN (" & w_Tipos & "))").Fields("Valor"), "R$ 0.00")
        
    End If

    If txt_VistT = "" Then txt_VistT = "R$ 0,00"
    If txt_NVistT = "" Then txt_NVistT = "R$ 0,00"
    If TXT_TOTALT = "" Then TXT_TOTALT = "R$ 0,00"
    
End If

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub ckZerados_Click()
    cmdPesq_Click
End Sub

Private Sub cmdPesq_Click()
Dim w_FirtTipo As Boolean
Dim w_Tipos
Dim w_tipo, J
Dim nenhumTipo As Boolean
Dim w_MaisVist, w_MaisNVist, w_MaisT As Double
Dim w_MenosVist, w_MenosNVist, w_MenosT As Double
Dim w_IgualVist, w_IgualNVist, w_IgualT As Double
Dim w_TotVist, w_TotNVist, w_TotT As Double


On Error GoTo err1

    If txt_tipo.Enabled = True Then
        nenhumTipo = True
        For J = 0 To txt_tipo.ListCount - 1
            If txt_tipo.Selected(J) = True Then
               nenhumTipo = False
            End If
        Next
        
        If nenhumTipo Then
                MsgBox "Selecione ao menos um TIPO!", vbCritical
                Exit Sub
        End If
    End If

    'tipos
    w_FirstTipo = True

    For J = 0 To txt_tipo.ListCount - 1
        If txt_tipo.Selected(J) = True Then
            Select Case txt_tipo.list(J)
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
        
            If w_FirstTipo Then
                w_Tipos = "'" & w_tipo & "'"
            Else
                w_Tipos = w_Tipos & ",'" & w_tipo & "'"
            End If
            w_FirstTipo = False
        End If
    Next

If ((TXT_LOGO = "" And ckTodas.value = 1) Or (TXT_LOGO <> "" And ckTodas.value = 0)) And ((dbNome = "" And ck_Nome.value = 1) Or (dbNome <> "" And ck_Nome.value = 0)) And ((TXT_CONTA = "" And ckConta.value = 1) Or (TXT_CONTA <> "" And ckConta.value = 0)) Then
    If ck_Nome.value = 0 Then
    'um nome
        If dbNome = "" Then
            wStrSql = "SELECT TAB_DESC_CALC.C_NCRED, TAB_DESC_CALC.c_codigo AS codigo, lojb010.num & ' ' & lojb010.cod_loj AS LOGO, TAB_FUNCIONARIO.F_NOME AS NOME, TAB_TP_CONTA.TP_DESC AS CONTA, TAB_DESC_CALC.C_DATA_INTERNA AS DATA, TAB_DESC_CALC.C_DESC AS DESCR, TAB_DESC_CALC.C_VALOR AS VALOR, TAB_DESC_CALC.C_TP_OP AS OP, TAB_DESC_CALC.C_VISTO AS VISTO, TAB_DESC_CALC.C_TP_CONTA, TAB_DESC_CALC.C_N_Ficha as Ficha, TAB_FICHA_MENS.M_MES, TAB_FICHA_MENS.M_ANO, TAB_FUNCIONARIO.F_CODIGO AS FUNC, TAB_DESC_CALC.C_DATA_INTERNA, TAB_TP_CONTA.TP_COD FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_TP_CONTA,"
            If ckFixos Then wStrSql = wStrSql & " TAB_DESC_CALC_FIXO,"
            wStrSql = wStrSql & " TAB_FUNCIONARIO INNER JOIN Lojb010 ON TAB_FUNCIONARIO.F_Cod_L = Lojb010.COD_LOJ WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " "
            If ckFixos Then wStrSql = wStrSql & "AND TAB_DESC_CALC.C_NCRED = TAB_DESC_CALC_FIXO.CF_CODIGO AND Month(TAB_DESC_CALC_FIXO.CF_DT) = " & TXT_MES & " AND Year(TAB_DESC_CALC_FIXO.CF_DT) = " & TXT_ANO
            If ckZerados = 0 Then wStrSql = wStrSql & "AND TAB_DESC_CALC.C_VALOR <> 0 "
            If Not acessoTotal() Then wStrSql = wStrSql & " AND ((TAB_DESC_CALC.C_TP_CONTA <> 20 AND TAB_DESC_CALC.C_TP_CONTA <> 78 and (F_TIPO <> 'V' AND F_TIPO <> 'C' AND F_TIPO <> 'X' AND F_TIPO <> '2')) OR (F_TIPO = 'V' OR F_TIPO = 'C' OR F_TIPO = 'X' OR F_TIPO = '2'))"
            wStrSql = wStrSql & " ORDER BY lojb010.num, TAB_FUNCIONARIO.F_NOME, TAB_DESC_CALC.C_TP_CONTA, TAB_DESC_CALC.C_DATA_INTERNA"
   
            'txt111.Text = wStrSql
   
        Else
         wStrSql = "SELECT TAB_DESC_CALC.C_NCRED, TAB_DESC_CALC.c_codigo AS codigo, lojb010.num & ' ' & lojb010.cod_loj AS LOGO, TAB_FUNCIONARIO.F_NOME AS NOME, TAB_TP_CONTA.TP_DESC AS CONTA, TAB_DESC_CALC.C_DATA_INTERNA AS DATA, TAB_DESC_CALC.C_DESC AS DESCR, TAB_DESC_CALC.C_VALOR AS VALOR, TAB_DESC_CALC.C_TP_OP AS OP, TAB_DESC_CALC.C_VISTO AS VISTO, TAB_DESC_CALC.C_TP_CONTA, TAB_DESC_CALC.C_N_Ficha as Ficha , TAB_FICHA_MENS.M_MES, TAB_FICHA_MENS.M_ANO, TAB_FUNCIONARIO.F_CODIGO AS FUNC, TAB_DESC_CALC.C_DATA_INTERNA, TAB_TP_CONTA.TP_COD FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_TP_CONTA,"
         If ckFixos Then wStrSql = wStrSql & " TAB_DESC_CALC_FIXO,"
         wStrSql = wStrSql & " TAB_FUNCIONARIO INNER JOIN Lojb010 ON TAB_FUNCIONARIO.F_Cod_L = Lojb010.COD_LOJ WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') AND (TAB_FICHA_MENS.M_NOME LIKE '" & dbNome & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " "
         If ckFixos Then wStrSql = wStrSql & "AND TAB_DESC_CALC.C_NCRED = TAB_DESC_CALC_FIXO.CF_CODIGO AND Month(TAB_DESC_CALC_FIXO.CF_DT) = " & TXT_MES & " AND Year(TAB_DESC_CALC_FIXO.CF_DT) = " & TXT_ANO
         If ckZerados = 0 Then wStrSql = wStrSql & "AND TAB_DESC_CALC.C_VALOR <> 0 "
         If Not acessoTotal() Then wStrSql = wStrSql & " AND ((TAB_DESC_CALC.C_TP_CONTA <> 20 AND TAB_DESC_CALC.C_TP_CONTA <> 78 and (F_TIPO <> 'V' AND F_TIPO <> 'C' AND F_TIPO <> 'X' AND F_TIPO <> '2')) OR (F_TIPO = 'V' OR F_TIPO = 'C' OR F_TIPO = 'X' OR F_TIPO = '2'))"
         wStrSql = wStrSql & " ORDER BY lojb010.num, TAB_FUNCIONARIO.F_NOME, TAB_DESC_CALC.C_TP_CONTA, TAB_DESC_CALC.C_DATA_INTERNA"
       
         'txt111.Text = wStrSql
       
            Set adoConta.Recordset = de.cnc.Execute(wStrSql).Clone
            
        End If
     
    
    Else
    'Todos os nomes
        wStrSql = "SELECT TAB_DESC_CALC.C_NCRED, TAB_DESC_CALC.c_codigo AS codigo, lojb010.num & ' ' & lojb010.cod_loj AS LOGO, TAB_FUNCIONARIO.F_NOME AS NOME, TAB_TP_CONTA.TP_DESC AS CONTA, TAB_DESC_CALC.C_DATA_INTERNA AS DATA, TAB_DESC_CALC.C_DESC AS DESCR, TAB_DESC_CALC.C_VALOR AS VALOR, TAB_DESC_CALC.C_TP_OP AS OP, TAB_DESC_CALC.C_VISTO AS VISTO, TAB_DESC_CALC.C_TP_CONTA, TAB_DESC_CALC.C_N_Ficha as Ficha, TAB_FICHA_MENS.M_MES, TAB_FICHA_MENS.M_ANO, TAB_FUNCIONARIO.F_CODIGO AS FUNC, TAB_DESC_CALC.C_DATA_INTERNA, TAB_TP_CONTA.TP_COD FROM TAB_FICHA_MENS, TAB_DESC_CALC, TAB_TP_CONTA,"
        If ckFixos Then wStrSql = wStrSql & " TAB_DESC_CALC_FIXO,"
        wStrSql = wStrSql & " TAB_FUNCIONARIO INNER JOIN Lojb010 ON TAB_FUNCIONARIO.F_Cod_L = Lojb010.COD_LOJ WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_logo LIKE '" & IIf(TXT_LOGO = "", "%", TXT_LOGO) & "') " & IIf(TXT_CONTA = "", "", " AND (TAB_TP_CONTA.TP_Cod = " & TXT_CONTA.BoundText & ")") & " AND ((TAB_FUNCIONARIO.F_TIPO) IN (" & w_Tipos & ")) "
        If ckFixos Then wStrSql = wStrSql & "AND TAB_DESC_CALC.C_NCRED = TAB_DESC_CALC_FIXO.CF_CODIGO AND Month(TAB_DESC_CALC_FIXO.CF_DT) = " & TXT_MES & " AND Year(TAB_DESC_CALC_FIXO.CF_DT) = " & TXT_ANO
        If ckZerados = 0 Then wStrSql = wStrSql & "AND TAB_DESC_CALC.C_VALOR <> 0 "
        If Not acessoTotal() Then wStrSql = wStrSql & " AND ((TAB_DESC_CALC.C_TP_CONTA <> 20 AND TAB_DESC_CALC.C_TP_CONTA <> 78 and (F_TIPO <> 'V' AND F_TIPO <> 'C' AND F_TIPO <> 'X' AND F_TIPO <> '2')) OR (F_TIPO = 'V' OR F_TIPO = 'C' OR F_TIPO = 'X' OR F_TIPO = '2'))"
        wStrSql = wStrSql & " ORDER BY lojb010.num, TAB_FUNCIONARIO.F_NOME, TAB_DESC_CALC.C_TP_CONTA, TAB_DESC_CALC.C_DATA_INTERNA"
        
        'txt111.Text = wStrSql
        
        Set adoConta.Recordset = de.cnc.Execute(wStrSql).Clone
 
    End If

        w_MaisVist = 0
    w_MaisNVist = 0
    w_MaisT = 0
    w_MenosVist = 0
    w_MenosNVist = 0
    w_MenosT = 0
    w_IgualVist = 0
    w_IgualNVist = 0
    w_IgualT = 0
    w_TotVist = 0
    w_TotNVist = 0
    w_TotT = 0
    
    If adoConta.Recordset.RecordCount <> 0 Then adoConta.Recordset.MoveFirst
    
    grid_Conta.Visible = False
    Do While Not adoConta.Recordset.EOF
        
        If adoConta.Recordset.Fields("visto") Then
            Select Case adoConta.Recordset.Fields("op")
                Case "+":
                    w_MaisVist = w_MaisVist + CDbl(adoConta.Recordset.Fields("Valor"))
                Case "-":
                    w_MenosVist = w_MenosVist + CDbl(adoConta.Recordset.Fields("Valor"))
                Case "=":
                    w_IgualVist = w_IgualVist + CDbl(adoConta.Recordset.Fields("Valor"))
            End Select
        Else
            Select Case adoConta.Recordset.Fields("op")
                Case "+":
                    w_MaisNVist = w_MaisNVist + CDbl(adoConta.Recordset.Fields("Valor"))
                Case "-":
                    w_MenosNVist = w_MenosNVist + CDbl(adoConta.Recordset.Fields("Valor"))
                Case "=":
                    w_IgualNVist = w_IgualNVist + CDbl(adoConta.Recordset.Fields("Valor"))
            End Select
        End If
        adoConta.Recordset.MoveNext
    Loop
    grid_Conta.Visible = True
    
    w_MaisT = w_MaisVist + w_MaisNVist
    w_MenosT = w_MenosVist + w_MenosNVist
    w_IgualT = w_IgualVist + w_IgualNVist
    
    w_TotVist = w_MaisVist + w_MenosVist
    w_TotNVist = w_MaisNVist + w_MenosNVist
    w_TotT = w_MaisT + w_MenosT
    
    txt_VistMais = Format(w_MaisVist, "R$ 0.00")
    txt_VistMenos = Format(w_MenosVist, "R$ 0.00")
    txt_VistIgual = Format(w_IgualVist, "R$ 0.00")
    txt_VistT = Format(w_TotVist, "R$ 0.00")
    
    txt_NVistMais = Format(w_MaisNVist, "R$ 0.00")
    txt_NVistMenos = Format(w_MenosNVist, "R$ 0.00")
    txt_NVistIgual = Format(w_IgualNVist, "R$ 0.00")
    txt_NVistT = Format(w_TotNVist, "R$ 0.00")
    
    TXT_TOTALMais = Format(w_MaisT, "R$ 0.00")
    TXT_TOTALMenos = Format(w_MenosT, "R$ 0.00")
    TXT_TOTALIgual = Format(w_IgualT, "R$ 0.00")
    TXT_TOTALT = Format(w_TotT, "R$ 0.00")
    
    If adoConta.Recordset.RecordCount <> 0 Then adoConta.Recordset.MoveFirst
    
    BarraF.Buttons("imprimir").Enabled = True
    grid_Conta.Enabled = True

    If adoConta.Recordset.RecordCount <= 0 Then MsgBox "Não existe fichas nestas características!", vbExclamation

Else
    ckConta_Click
    'MsgBox "Preencha os dados p/ consulta!", vbCritical
    'TXT_CONTA_cod.SetFocus
End If


sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub



Private Sub dbNome_Change()
'    If dbNome.text = "%" Then
'        ckTodas.value = 0
'        ckConta.value = 0
'    Else
'        If w_load Then ckTodas.value = 1
'        ckConta.value = 1
'        cmdPesq_Click
'    End If
    If (dbNome.text <> "" And dbNome.text <> "%") Then
        If (ck_Nome.value = 0) Then
            ckConta.value = 1
            ckConta_Click
            ckTodas.value = 1
            ckTodas_Click
        End If
    End If
        
End Sub

Private Sub dbNome_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then Sendkeys "{tab}"

End Sub

Private Sub Form_Activate()
    w_load = False
    
    For J = 0 To txt_tipo.ListCount - 1
        If txt_tipo.list(J) = frm_Alt_Fic_Mensal_VIS.TXT_FTIPO.Caption Then
           txt_tipo.Selected(J) = True
        End If
    Next
    
    If ck_Nome.value = 1 Then
        txt_tipo.Enabled = True
        ckTipo.Enabled = True
    Else
        txt_tipo.Enabled = False
        ckTipo.Enabled = False
    End If
    
    ckTipo.value = 1
    ckTipo_Click
    ckConta.value = 1
    ckConta_Click
    ck_Nome.value = 1
    ck_Nome_Click
    ckTodas.value = 1
    ckTodas_Click
    
    w_load = True
    
    cmdPesq_Click
    
    ckConta.value = 0
    TXT_CONTA_cod.SetFocus
    
End Sub

Private Sub Form_Load()
On Error GoTo err1
    
    TXT_MES = Format(Date, "mm")
    TXT_ANO = Format(Date, "yyyy")
    Set ADO_FUNC.Recordset = de.cnc.Execute("SELECT F_Nome, F_CPF FROM TAB_FUNCIONARIO Where NOT F_NOME = '10 - Func' and not F_NOME = '99 - Presence' GROUP BY F_Nome, F_CPF ORDER BY F_NOME").Clone
    
    If (UCase(frmLogin.txtUserName.text) = NomeMestre Or UCase(frmLogin.txtUserName.text) = NomeMestre2 Or UCase(frmLogin.txtUserName.text) = NomeMestre3) Then
        mnu.Visible = True
        'Height = 8355
    End If
    
    
sair:
    Exit Sub
    
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
    
End Sub



Private Sub grid_Conta_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 2 And grid_Conta.Enabled = True And (UCase(frmLogin.txtUserName.text) = NomeMestre Or UCase(frmLogin.txtUserName.text) = NomeMestre2 Or UCase(frmLogin.txtUserName.text) = NomeMestre3) Then PopupMenu mnu
End Sub



'--------- Ao Pressionar uma Tecla -----------

Private Sub grid_Conta_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub GRID_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub





Private Sub mnutodosR_Click()
On erro GoTo err1

If (UCase(w_usuario) <> UCase(NomeMestre) And UCase(w_usuario) <> UCase(NomeMestre2) And UCase(w_usuario) <> UCase(NomeMestre3)) Then Exit Sub


    W_POS = adoConta.Recordset.AbsolutePosition

    adoConta.Recordset.MoveFirst
    Do While Not adoConta.Recordset.EOF
    
            'Atualiza Visto
            w_cod = adoConta.Recordset.Fields("Codigo")
            W_NFICHA = adoConta.Recordset.Fields("FICHA")
            W_F_COD = adoConta.Recordset.Fields("FUNC")
            
            w_data = adoConta.Recordset.Fields("DATA")
            
            'If isMesValido(W_F_COD, Month(w_data), Year(w_data)) Then 'Verifica se é mês atual ou passado
            
                    '*** ATUALIZA TAB_FUNCIONARIO O CAMPO OK   SE   FOR   FERIAS OU 13ºSALARIO
                    If adoConta.Recordset.Fields("C_TP_CONTA") = "24" Then      'FERIAS
                        de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS_OK = 0  WHERE (F_CODIGO = " & W_F_COD & " )", w_reg
                        de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_FERIAS_OK = 0  WHERE (M_NFICHA = " & W_NFICHA & " )", w_reg
                    
                    ElseIf adoConta.Recordset.Fields("C_TP_CONTA") = "32" Then   '*** 13º
                        de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_13_OK = 0  WHERE (F_CODIGO = " & W_F_COD & " )", w_reg
                        de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_13_OK = 0  WHERE (M_NFICHA = " & W_NFICHA & " )", w_reg
                    End If
            
                'Atualiza Visto
                de.cnc.Execute "Update TAB_DESC_CALC Set C_VISTO = 0 Where (C_CODIGO = " & w_cod & ")"
            
            'Else
                'MsgBox "O lançamento '" & w_cod & "' não pode ser alterado por não ser do mês atual!", vbCritical, "Erro!"
            'End If
                adoConta.Recordset.MoveNext
    Loop
    
    
    cmdPesq_Click
    adoConta.Recordset.Move W_POS - 1

sair:
    Exit Sub
err1:
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub mnutodos_Click()
On erro GoTo err1


If (UCase(w_usuario) <> UCase(NomeMestre) And UCase(w_usuario) <> UCase(NomeMestre2) And UCase(w_usuario) <> UCase(NomeMestre3)) Then Exit Sub
    
    W_POS = adoConta.Recordset.AbsolutePosition


    adoConta.Recordset.MoveFirst
    Do While Not adoConta.Recordset.EOF
                
            'Atualiza Visto
            w_cod = adoConta.Recordset.Fields("Codigo")
            W_NFICHA = adoConta.Recordset.Fields("FICHA")
            W_F_COD = adoConta.Recordset.Fields("FUNC")
            
            w_data = adoConta.Recordset.Fields("DATA")
            
            'If isMesValido(W_F_COD, Month(w_data), Year(w_data)) Then 'Verifica se é mês atual ou passado
                'If CDbl(adoConta.Recordset.Fields("VALOR")) <> 0 Then
                    '*** ATUALIZA TAB_FUNCIONARIO O CAMPO OK   SE   FOR   FERIAS OU 13ºSALARIO
                    If adoConta.Recordset.Fields("C_TP_CONTA") = "24" Then      'FERIAS
                        de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS_OK = -1  WHERE (F_CODIGO = " & W_F_COD & " )", w_reg
                        de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_FERIAS_OK = -1  WHERE (M_NFICHA = " & W_NFICHA & " )", w_reg
                    
                    ElseIf adoConta.Recordset.Fields("C_TP_CONTA") = "32" Then   '*** 13º
                        de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_13_OK = -1  WHERE (F_CODIGO = " & W_F_COD & " )", w_reg
                        de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_13_OK = -1  WHERE (M_NFICHA = " & W_NFICHA & " )", w_reg
                    End If
                
                    'Atualiza Visto
                    de.cnc.Execute "Update TAB_DESC_CALC Set C_VISTO = -1 Where (C_CODIGO = " & w_cod & ")"
                'End If
            'Else
                'MsgBox "O lançamento '" & w_cod & "' não pode ser alterado por não ser do mês atual!", vbCritical, "Erro!"
            'End If
                
        adoConta.Recordset.MoveNext
    Loop
    
    
    cmdPesq_Click
    adoConta.Recordset.Move W_POS - 1
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub mnuVist_Click()
On erro GoTo err1
    
    
If (UCase(w_usuario) <> UCase(NomeMestre) And UCase(w_usuario) <> UCase(NomeMestre2) And UCase(w_usuario) <> UCase(NomeMestre3)) Then Exit Sub

    
        'Atualiza Visto
    w_cod = adoConta.Recordset.Fields("Codigo")
    W_NFICHA = adoConta.Recordset.Fields("FICHA")
    W_F_COD = adoConta.Recordset.Fields("FUNC")
    
    w_data = adoConta.Recordset.Fields("DATA")
    
    'If isMesValido(W_F_COD, Month(w_data), Year(w_data)) Then 'Verifica se é mês atual ou passado
    
    
        '*** ATUALIZA TAB_FUNCIONARIO O CAMPO OK   SE   FOR   FERIAS OU 13ºSALARIO
        If adoConta.Recordset.Fields("C_TP_CONTA") = "24" Or adoConta.Recordset.Fields("C_TP_CONTA") = "68" Then      'FERIAS
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS_OK =-1  WHERE (F_CODIGO = " & W_F_COD & " )", w_reg
            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_FERIAS_OK = -1  WHERE (M_NFICHA = " & W_NFICHA & " )", w_reg
        
        ElseIf adoConta.Recordset.Fields("C_TP_CONTA") = "32" Then   '*** 13º
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_13_OK = -1  WHERE (F_CODIGO = " & W_F_COD & " )", w_reg
            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_13_OK = -1  WHERE (M_NFICHA = " & W_NFICHA & " )", w_reg
        End If
        
        
        'Atualiza Visto
        de.cnc.Execute "Update TAB_DESC_CALC Set C_VISTO = -1 Where (C_CODIGO = " & w_cod & ")"
                
        W_POS = adoConta.Recordset.AbsolutePosition
        cmdPesq_Click
        adoConta.Recordset.Move W_POS - 1
        
        
   ' Else
    '    MsgBox "O lançamento '" & w_cod & "' não pode ser alterado por não ser do mês atual!", vbCritical, "Erro!"
    'End If
        

    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub mnuVolt_Click()
On erro GoTo err1
    
    
If (UCase(w_usuario) <> UCase(NomeMestre) And UCase(w_usuario) <> UCase(NomeMestre2) And UCase(w_usuario) <> UCase(NomeMestre3)) Then Exit Sub


        'Atualiza Visto
    w_cod = adoConta.Recordset.Fields("Codigo")
    W_NFICHA = adoConta.Recordset.Fields("FICHA")
    W_F_COD = adoConta.Recordset.Fields("FUNC")
    
    w_data = adoConta.Recordset.Fields("DATA")
    
    'If isMesValido(W_F_COD, Month(w_data), Year(w_data)) Then 'Verifica se é mês atual ou passado
    
        '*** ATUALIZA TAB_FUNCIONARIO O CAMPO OK   SE   FOR   FERIAS OU 13ºSALARIO
        If adoConta.Recordset.Fields("C_TP_CONTA") = "24" Then      'FERIAS
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS_OK = 0  WHERE (F_CODIGO = " & W_F_COD & " )", w_reg
            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_FERIAS_OK = 0  WHERE (M_NFICHA = " & W_NFICHA & " )", w_reg
        
        ElseIf adoConta.Recordset.Fields("C_TP_CONTA") = "32" Then   '*** 13º
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_13_OK = 0  WHERE (F_CODIGO = " & W_F_COD & " )", w_reg
            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_13_OK = 0  WHERE (M_NFICHA = " & W_NFICHA & " )", w_reg
        End If
        
        'Atualiza Não Vistado
        de.cnc.Execute "Update TAB_DESC_CALC Set C_VISTO = 0 Where (C_CODIGO = " & w_cod & ")"
        W_POS = adoConta.Recordset.AbsolutePosition
        cmdPesq_Click
        adoConta.Recordset.Move W_POS - 1
    
    'Else
    '    MsgBox "O lançamento '" & w_cod & "' não pode ser alterado por não ser do mês atual!", vbCritical, "Erro!"
    'End If

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair

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

Sub TXT_CONTA_cod_LostFocus()
    If TXT_CONTA_cod <> "" Then
        If TXT_CONTA_cod <> "" Then
            TXT_CONTA.BoundText = Int(TXT_CONTA_cod)
        Else
            ckConta_Click
            Exit Sub
        End If
    End If
    cmdPesq_Click
End Sub

Private Sub TXT_CONTA_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 And ckTodas.value = 0 Then
        TXT_LOGO2.SetFocus
        cmdPesq_Click
      ElseIf KeyCode = 13 Then
        ckTodas.SetFocus
        cmdPesq_Click
      End If
End Sub

Private Sub TXT_LOGO_Change()
   'If TXT_LOGO <> "" Then ck_Nome.value = 1
   'TXT_LOGO2.BoundText = TXT_LOGO.BoundText
  
   'If TXT_CONTA_cod.Text <> "" Then cmdPesq_Click
   
   

   
End Sub

Private Sub TXT_LOGO_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then
      
         If TXT_LOGO <> "" Then ck_Nome.value = 1
         TXT_LOGO2.BoundText = TXT_LOGO.BoundText
        
         If TXT_CONTA_cod.text <> "" Then
            Sendkeys "{tab}"
            cmdPesq_Click
        End If
        
      End If
End Sub

Private Sub TXT_LOGO_Validate(Cancel As Boolean)
    If TXT_LOGO <> "" Then ck_Nome.value = 1
         TXT_LOGO2.BoundText = TXT_LOGO.BoundText
        
    If TXT_CONTA_cod.text <> "" Then
        Sendkeys "{tab}"
        cmdPesq_Click
    End If
End Sub

Private Sub TXT_LOGO2_Change()
    'TXT_LOGO.BoundText = TXT_LOGO2.BoundText

End Sub

Private Sub TXT_LOGO2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
      TXT_LOGO.BoundText = TXT_LOGO2.BoundText
      Sendkeys "{tab}"
      cmdPesq_Click
    End If

End Sub

Private Sub TXT_LOGO2_LostFocus()
      TXT_LOGO.BoundText = TXT_LOGO2.BoundText
      Sendkeys "{tab}"
      cmdPesq_Click
End Sub

Private Sub TXT_LOGO2_Validate(Cancel As Boolean)
      TXT_LOGO.BoundText = TXT_LOGO2.BoundText
      Sendkeys "{tab}"
      cmdPesq_Click
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
    Case 73: ' "I"
            Imprimir
    End Select
End If
End Sub
