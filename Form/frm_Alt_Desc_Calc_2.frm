VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.6#0"; "ACTIVETEXT.ocx"
Begin VB.Form frm_Alt_Desc_Calc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALTERAÇÃO E CADASTRO DE CONTAS"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13005
   Icon            =   "frm_Alt_Desc_Calc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   13005
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4680
      Top             =   960
   End
   Begin MSDataGridLib.DataGrid Grid 
      Bindings        =   "frm_Alt_Desc_Calc.frx":1042
      Height          =   2415
      Left            =   6360
      TabIndex        =   77
      Top             =   3360
      Visible         =   0   'False
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   4260
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "DESCRIÇÃO DE CÁLCULOS"
      ColumnCount     =   2
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "TP"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   2459,906
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab GUIA 
      Height          =   3975
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "Alteração"
      TabPicture(0)   =   "frm_Alt_Desc_Calc.frx":1059
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label13"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LBLOGO"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "LBNCRED"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbEmp(5)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbEmp(6)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbEmp(7)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "LB_EMP_DE(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "LB_EMP_DE(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lb_OBS"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lb_dt_13"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txt_13"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TXT_CONTA_cod"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txt_DT"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "TXT_CONTA"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TXT_DESC"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TXT_OP"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txt_nficha"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txt_conta_Op"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "TXT_LOGO"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "TXT_NUM"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "ADO_CRED"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txt_valor"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txt_Emp(4)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txt_Emp(5)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "TXT_E_COD_E"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "TXT_E_JUROS_E"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txt_Obs"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txt_ncred"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "chkVisto"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).ControlCount=   34
      TabCaption(1)   =   "Cadastrar"
      TabPicture(1)   =   "frm_Alt_Desc_Calc.frx":1075
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TXT_DESC_EXTRA"
      Tab(1).Control(1)=   "TXT_DT_EXTRA"
      Tab(1).Control(2)=   "txt_Emp(0)"
      Tab(1).Control(3)=   "txt_Emp(2)"
      Tab(1).Control(4)=   "TXT_E_SALDO"
      Tab(1).Control(5)=   "TXT_E_VALOR"
      Tab(1).Control(6)=   "TXT_E_COD"
      Tab(1).Control(7)=   "txt_Emp(1)"
      Tab(1).Control(8)=   "txt_NCred_Cad"
      Tab(1).Control(9)=   "txt_Logo_Cad"
      Tab(1).Control(10)=   "TXT_CONTA_CAD_op"
      Tab(1).Control(11)=   "TXT_NFICHA_CAD"
      Tab(1).Control(12)=   "TXT_DT_CAD"
      Tab(1).Control(13)=   "TXT_CONTA_CAD"
      Tab(1).Control(14)=   "TXT_DESC_CAD"
      Tab(1).Control(15)=   "TXT_OP_CAD"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "TXT_CONTA_Cod_CAD"
      Tab(1).Control(17)=   "TXT_VALOR_CAD"
      Tab(1).Control(18)=   "TXT_E_JUROS"
      Tab(1).Control(19)=   "LB_DESC_EXTRA"
      Tab(1).Control(20)=   "LB_DT_EXTRA"
      Tab(1).Control(21)=   "LB_EMP_D(3)"
      Tab(1).Control(22)=   "lbEmp(3)"
      Tab(1).Control(23)=   "LB_EMP_D(2)"
      Tab(1).Control(24)=   "LB_EMP_D(1)"
      Tab(1).Control(25)=   "LB_EMP_D(0)"
      Tab(1).Control(26)=   "lbEmp(0)"
      Tab(1).Control(27)=   "lbEmp(1)"
      Tab(1).Control(28)=   "lbEmp(2)"
      Tab(1).Control(29)=   "lbncred_cad"
      Tab(1).Control(30)=   "lblogo_cad"
      Tab(1).Control(31)=   "Label8"
      Tab(1).Control(32)=   "Label12"
      Tab(1).Control(33)=   "Label11"
      Tab(1).Control(34)=   "Label10"
      Tab(1).Control(35)=   "Label7"
      Tab(1).Control(36)=   "Label6"
      Tab(1).Control(37)=   "Shape2"
      Tab(1).ControlCount=   38
      Begin VB.CheckBox chkVisto 
         Caption         =   "VISTO"
         DataField       =   "C_VISTO"
         DataSource      =   "ADOREG"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4080
         TabIndex        =   76
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txt_ncred 
         Alignment       =   2  'Center
         DataField       =   "C_NCRED"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
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
         Left            =   4080
         TabIndex        =   75
         Top             =   2640
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txt_Obs 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   71
         Top             =   3120
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.TextBox TXT_DESC_EXTRA 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   -73680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   69
         Top             =   3210
         Width           =   3525
      End
      Begin rdActiveText.ActiveText TXT_DT_EXTRA 
         Height          =   345
         Left            =   -74760
         TabIndex        =   67
         Top             =   3330
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
         TextMask        =   1
         RawText         =   1
         Mask            =   "##/##/####"
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.TextBox TXT_E_JUROS_E 
         Alignment       =   2  'Center
         DataField       =   "C_JUROS"
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
         Left            =   1440
         TabIndex        =   65
         Top             =   3360
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.TextBox txt_Emp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   0
         Left            =   -71445
         TabIndex        =   15
         Top             =   3360
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox txt_Emp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   2
         Left            =   -72240
         TabIndex        =   14
         Top             =   3360
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox TXT_E_COD_E 
         Alignment       =   2  'Center
         DataField       =   "C_EMP_COD"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
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
         Left            =   225
         TabIndex        =   60
         Top             =   3360
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox TXT_E_SALDO 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
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
         Left            =   -72450
         TabIndex        =   56
         Top             =   3360
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox TXT_E_VALOR 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
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
         Left            =   -73710
         TabIndex        =   55
         Top             =   3360
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox TXT_E_COD 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
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
         Left            =   -74760
         TabIndex        =   54
         Top             =   3360
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txt_Emp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         DataField       =   "C_QTDE_PARC"
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
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   5
         Left            =   3480
         TabIndex        =   46
         Top             =   3375
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox txt_Emp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         CausesValidation=   0   'False
         DataField       =   "C_JUROS"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0,00"
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
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   4
         Left            =   4215
         TabIndex        =   47
         Top             =   3360
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox txt_Emp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   1
         Left            =   -70680
         TabIndex        =   16
         Top             =   3360
         Visible         =   0   'False
         Width           =   540
      End
      Begin rdActiveText.ActiveText txt_valor 
         CausesValidation=   0   'False
         DataField       =   "C_VALOR"
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
         Height          =   375
         Left            =   3510
         TabIndex        =   2
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   13
         Text            =   "R$ 0,00"
         FocusSelect     =   0   'False
         RawText         =   0
         FloatFormat     =   2
         eAuto           =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.TextBox txt_NCred_Cad 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
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
         Left            =   -74160
         TabIndex        =   41
         Top             =   3360
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txt_Logo_Cad 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
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
         Left            =   -74760
         TabIndex        =   39
         Top             =   3360
         Visible         =   0   'False
         Width           =   585
      End
      Begin MSAdodcLib.Adodc ADO_CRED 
         Height          =   330
         Left            =   2400
         Top             =   3600
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   4
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
         Caption         =   "CRED"
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
      Begin VB.TextBox TXT_NUM 
         Alignment       =   1  'Right Justify
         DataField       =   "C_NCRED"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
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
         Left            =   840
         TabIndex        =   36
         Top             =   3360
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox TXT_LOGO 
         Alignment       =   1  'Right Justify
         DataField       =   "C_LOGO"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
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
         Left            =   195
         TabIndex        =   34
         Top             =   3360
         Visible         =   0   'False
         Width           =   585
      End
      Begin MSDataListLib.DataCombo TXT_CONTA_CAD_op 
         Bindings        =   "frm_Alt_Desc_Calc.frx":1091
         DataSource      =   "ADOREG"
         Height          =   360
         Left            =   -71160
         TabIndex        =   33
         Top             =   1365
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "TP_OP"
         BoundColumn     =   "TP_COD"
         Text            =   ""
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
      Begin MSDataListLib.DataCombo txt_conta_Op 
         Bindings        =   "frm_Alt_Desc_Calc.frx":10A2
         DataField       =   "C_TP_CONTA"
         DataSource      =   "ADOREG"
         Height          =   360
         Left            =   3720
         TabIndex        =   32
         Top             =   1320
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "TP_OP"
         BoundColumn     =   "TP_COD"
         Text            =   ""
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
      Begin VB.TextBox txt_nficha 
         Alignment       =   1  'Right Justify
         DataField       =   "c_N_ficha"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
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
         Left            =   195
         TabIndex        =   0
         Top             =   600
         Width           =   1305
      End
      Begin VB.TextBox TXT_NFICHA_CAD 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
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
         Left            =   -74805
         TabIndex        =   7
         Top             =   675
         Width           =   1305
      End
      Begin VB.ComboBox TXT_OP 
         DataField       =   "C_TP_OP"
         DataSource      =   "ADOREG"
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
         ItemData        =   "frm_Alt_Desc_Calc.frx":10B3
         Left            =   4275
         List            =   "frm_Alt_Desc_Calc.frx":10C0
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "+"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox TXT_DESC 
         DataField       =   "C_DESC"
         DataSource      =   "ADOREG"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   195
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2040
         Width           =   3855
      End
      Begin MSDataListLib.DataCombo TXT_CONTA 
         Bindings        =   "frm_Alt_Desc_Calc.frx":10CD
         DataField       =   "C_TP_CONTA"
         DataSource      =   "ADOREG"
         Height          =   360
         Left            =   915
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1320
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "TP_DESC"
         BoundColumn     =   "TP_COD"
         Text            =   ""
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
      Begin MSComCtl2.DTPicker txt_DT 
         DataField       =   "C_DT"
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
         Left            =   1965
         TabIndex        =   1
         Top             =   615
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   609
         _Version        =   393216
         Format          =   68550657
         CurrentDate     =   38432
      End
      Begin MSComCtl2.DTPicker TXT_DT_CAD 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   345
         Left            =   -73110
         TabIndex        =   8
         Top             =   690
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   609
         _Version        =   393216
         Format          =   68550657
         CurrentDate     =   38432
      End
      Begin MSDataListLib.DataCombo TXT_CONTA_CAD 
         Bindings        =   "frm_Alt_Desc_Calc.frx":10DE
         Height          =   360
         Left            =   -74085
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1350
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "TP_DESC"
         BoundColumn     =   "TP_COD"
         Text            =   ""
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
      Begin VB.TextBox TXT_DESC_CAD 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   -74805
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   2040
         Width           =   3855
      End
      Begin VB.ComboBox TXT_OP_CAD 
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
         ItemData        =   "frm_Alt_Desc_Calc.frx":10EF
         Left            =   -70725
         List            =   "frm_Alt_Desc_Calc.frx":10FC
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "+"
         Top             =   2070
         Width           =   615
      End
      Begin MSDataListLib.DataCombo TXT_CONTA_cod 
         Bindings        =   "frm_Alt_Desc_Calc.frx":1109
         DataField       =   "C_TP_CONTA"
         DataSource      =   "ADOREG"
         Height          =   360
         Left            =   195
         TabIndex        =   3
         Top             =   1320
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         ListField       =   "TP_COD"
         BoundColumn     =   "TP_COD"
         Text            =   ""
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
      Begin MSDataListLib.DataCombo TXT_CONTA_Cod_CAD 
         Bindings        =   "frm_Alt_Desc_Calc.frx":111A
         Height          =   360
         Left            =   -74805
         TabIndex        =   10
         Top             =   1350
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         ListField       =   "TP_COD"
         BoundColumn     =   "TP_COD"
         Text            =   ""
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
      Begin rdActiveText.ActiveText TXT_VALOR_CAD 
         Height          =   375
         Left            =   -71325
         TabIndex        =   9
         Top             =   690
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   13
         Text            =   "R$ 0,00"
         TextMask        =   4
         RawText         =   4
         FloatFormat     =   2
         eAuto           =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.TextBox TXT_E_JUROS 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
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
         Left            =   -71160
         TabIndex        =   62
         Top             =   3360
         Visible         =   0   'False
         Width           =   1020
      End
      Begin rdActiveText.ActiveText txt_13 
         Height          =   345
         Left            =   240
         TabIndex        =   73
         Top             =   3210
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
         TextMask        =   1
         RawText         =   1
         Mask            =   "##/##/####"
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.Label lb_dt_13 
         BackStyle       =   0  'Transparent
         Caption         =   "DATA (13º)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lb_OBS 
         BackStyle       =   0  'Transparent
         Caption         =   "OBS 13º OU TXT FÉRIAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1500
         TabIndex        =   72
         Top             =   3075
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label LB_DESC_EXTRA 
         BackStyle       =   0  'Transparent
         Caption         =   "OBS 13º OU TXT FÉRIAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -73575
         TabIndex        =   70
         Top             =   3000
         Width           =   3255
      End
      Begin VB.Label LB_DT_EXTRA 
         BackStyle       =   0  'Transparent
         Caption         =   "DATA (F)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   68
         Top             =   3120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label LB_EMP_DE 
         BackStyle       =   0  'Transparent
         Caption         =   "EMP. JUROS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   66
         Top             =   3120
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label LB_EMP_D 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "V.JUROS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   -71160
         TabIndex        =   63
         Top             =   3120
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lbEmp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIA Pg."
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   3
         Left            =   -72180
         TabIndex        =   64
         Top             =   3120
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label LB_EMP_DE 
         BackStyle       =   0  'Transparent
         Caption         =   "EMP. COD."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   61
         Top             =   3120
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label LB_EMP_D 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SALDO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   -72450
         TabIndex        =   59
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LB_EMP_D 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "VALOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   -73710
         TabIndex        =   58
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LB_EMP_D 
         BackStyle       =   0  'Transparent
         Caption         =   "COD."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   57
         Top             =   3120
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lbEmp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde de"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   7
         Left            =   3525
         TabIndex        =   51
         Top             =   3000
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lbEmp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   6
         Left            =   4440
         TabIndex        =   50
         Top             =   3135
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label lbEmp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parcelas"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   5
         Left            =   3495
         TabIndex        =   49
         Top             =   3195
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lbEmp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parcelas"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   0
         Left            =   -71430
         TabIndex        =   48
         Top             =   3180
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lbEmp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   -70470
         TabIndex        =   45
         Top             =   3120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label lbEmp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde de"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   2
         Left            =   -71415
         TabIndex        =   44
         Top             =   2985
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lbncred_cad 
         BackStyle       =   0  'Transparent
         Caption         =   "NUM."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74040
         TabIndex        =   42
         Top             =   3120
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblogo_cad 
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   40
         Top             =   3120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label LBNCRED 
         BackStyle       =   0  'Transparent
         Caption         =   "NUM."
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
         Left            =   840
         TabIndex        =   37
         Top             =   3120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label LBLOGO 
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
         Left            =   195
         TabIndex        =   35
         Top             =   3120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label13 
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
         Left            =   195
         TabIndex        =   30
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74820
         TabIndex        =   29
         Top             =   435
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "VALOR"
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
         Left            =   3795
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "DATA"
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
         Left            =   1965
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CONTA"
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
         Left            =   195
         TabIndex        =   21
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIÇÃO"
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
         Left            =   195
         TabIndex        =   20
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "OP."
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
         Left            =   4275
         TabIndex        =   19
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "VALOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -71220
         TabIndex        =   28
         Top             =   435
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "DATA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -73110
         TabIndex        =   27
         Top             =   435
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "CONTA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74805
         TabIndex        =   26
         Top             =   1110
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIÇÃO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74805
         TabIndex        =   25
         Top             =   1830
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "OP."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -70725
         TabIndex        =   24
         Top             =   1830
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   3495
         Left            =   -74925
         Top             =   390
         Width           =   4905
      End
   End
   Begin VB.CheckBox c_Filtro_Emp 
      Caption         =   "Mostrar todos (Empréstimos)"
      Height          =   225
      Left            =   1560
      TabIndex        =   53
      Top             =   5175
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid GRID_EMP 
      Bindings        =   "frm_Alt_Desc_Calc.frx":112B
      Height          =   1755
      Left            =   60
      TabIndex        =   52
      Top             =   5460
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3096
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
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
      Caption         =   "EMPRÉSTIMOS"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "E_CODIGO"
         Caption         =   "CODIGO"
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
         DataField       =   "E_DT_ULT_PG"
         Caption         =   "   DT. ÚLT. PG."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "E_VALOR"
         Caption         =   "  VALOR EMP."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "R$ 0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "E_SALDO"
         Caption         =   "           SALDO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "R$ 0.00"
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
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1154,835
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ado_EMP 
      Height          =   330
      Left            =   2760
      Top             =   6720
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Emprestimo"
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
   Begin MSDataGridLib.DataGrid GRID_CRED 
      Bindings        =   "frm_Alt_Desc_Calc.frx":1141
      Height          =   1755
      Left            =   60
      TabIndex        =   43
      Top             =   5460
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3096
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   0   'False
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
      Caption         =   "CRED.s"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "CRED"
         Caption         =   "CONTRATO"
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
         DataField       =   "VCTO"
         Caption         =   "   DT. VCTO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "VALOR"
         Caption         =   "           VALOR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "R$ 0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "SALDO"
         Caption         =   "           SALDO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "R$ 0.00"
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
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1154,835
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox c_Filtro 
      Caption         =   "Mostrar todos (crediários)"
      Height          =   225
      Left            =   1440
      TabIndex        =   38
      Top             =   5175
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8040
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Desc_Calc.frx":1158
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Desc_Calc.frx":1472
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Desc_Calc.frx":178C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Desc_Calc.frx":1AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Desc_Calc.frx":1DC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Desc_Calc.frx":20DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Desc_Calc.frx":23F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Desc_Calc.frx":452E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Desc_Calc.frx":4E08
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   1429
      ButtonWidth     =   1535
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
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
            Caption         =   "&Adicionar"
            Key             =   "adicionar"
            Object.ToolTipText     =   "Adicionar (Alt+A)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Editar"
            Key             =   "editar"
            Object.ToolTipText     =   "Editar Alteração (Alt+E)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Salvar"
            Key             =   "salvar"
            Object.ToolTipText     =   "Salvar Alteração (Alt+S)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Cancelar"
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar Alteração (Alt+C)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E&xcluir"
            Key             =   "excluir"
            Object.ToolTipText     =   "Excluir registro (Alt+X)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADOREG 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   7260
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   582
      ConnectMode     =   0
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
   Begin MSAdodcLib.Adodc ADO_GRID 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6930
      Visible         =   0   'False
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   582
      ConnectMode     =   0
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexGRID 
      Bindings        =   "frm_Alt_Desc_Calc.frx":6B12
      Height          =   6255
      Left            =   5280
      TabIndex        =   78
      Top             =   960
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   11033
      _Version        =   393216
      Rows            =   7
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      GridLineWidthFixed=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
   End
   Begin VB.Label lb_form 
      Caption         =   "Label3"
      Height          =   375
      Left            =   7440
      TabIndex        =   79
      Top             =   3720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label LB_FUNC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FC"
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
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   870
      Width           =   4935
   End
   Begin VB.Menu mnuSel 
      Caption         =   "select"
      Visible         =   0   'False
      Begin VB.Menu mnuSelSel 
         Caption         =   "Selecionar"
      End
   End
End
Attribute VB_Name = "frm_Alt_Desc_Calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W_LD_FILTRO As Boolean
Dim V_MOVE As Boolean
Dim V_MOVE_GRID As Boolean
Dim v_filtro As String
Dim v_filtro_puro As String
Dim w_At As Boolean
Dim w_PSS As String
Dim w_txt_desc As String
Dim w_Editar As Boolean

Sub formatarFlexGrid()
    If Not (adoReg.Recordset.EOF) Then

        For I = 1 To flexGRID.Rows - 1
            If flexGRID.TextMatrix(I, 10) > 0 Then
                 ' Varre todas as colunas da linha e seta a cor de fundo
                 For coluna = 0 To flexGRID.Cols - 1
                     flexGRID.Col = coluna
                     flexGRID.Row = I
                     flexGRID.CellBackColor = vbYellow
                     flexGRID.CellFontBold = True
                     'flexGRID_L.CellForeColor = &H80000004
                 Next coluna
             End If
        Next I
        
        flexGRID.ColAlignment(4) = flexAlignRightBottom 'valor
        flexGRID.ColAlignment(5) = flexAlignCenterBottom 'op
    
    End If

End Sub


Private Sub c_Filtro_Click()
On Error GoTo err1

    If c_Filtro.value = 1 Then
        'ADO_CRED.Recordset.Filter = v_filtro_puro
        ADO_CRED.Recordset.Filter = v_filtro_puro
    ElseIf c_Filtro.value = 0 Then
'        ADO_CRED.Recordset.Filter = v_filtro
        ADO_CRED.Recordset.Filter = v_filtro
    End If

    ADO_CRED.Recordset.Sort = "vcto"

sair:
    Exit Sub
err1:
    If Not Err.Number = 3705 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub



Private Sub c_Filtro_Emp_Click()
    If c_Filtro_Emp = 0 Then
        If TXT_E_COD <> "" Then ado_EMP.Recordset.Filter = "E_CODIGO = " & CDbl(TXT_E_COD)
    
    Else
        ado_EMP.Recordset.Filter = 0
        ado_EMP.Refresh
    
    End If
End Sub


Private Sub flexGRID_RowColChange()
    If flexGRID.RowSel <> 0 Then
        ADO_GRID.Recordset.MoveFirst
        ADO_GRID.Recordset.Find "C_CODIGO = " & flexGRID.TextMatrix(flexGRID.RowSel, 7), , adSearchForward
    End If
End Sub

Private Sub Form_Activate()
If w_At = True Then adoReg.Refresh
    
End Sub

Private Sub Form_Load()
On Error GoTo err1

flexGRID.ColWidth(0) = 900 'data
flexGRID.ColWidth(1) = 420 'COD
flexGRID.ColWidth(2) = 1600 'TP conta
flexGRID.ColWidth(3) = 2500 'conta
flexGRID.ColWidth(4) = 1080 'valor
flexGRID.ColWidth(5) = 330 'op
flexGRID.ColWidth(6) = 550 'visto
flexGRID.ColWidth(7) = 0
flexGRID.ColWidth(8) = 0
flexGRID.ColWidth(9) = 0
flexGRID.ColWidth(10) = 0 'ncred
flexGRID.ColWidth(11) = 0
flexGRID.ColWidth(12) = 0
flexGRID.ColWidth(13) = 0
flexGRID.ColWidth(14) = 0
flexGRID.ColWidth(15) = 0


'LIMPA AS CX. DE CADASTRO
    w_At = True
        
    If frm_Alt_Fic_Mensal_VIS.TXT_FUNC = "" Then
        TXT_NFICHA_CAD = frm_Alt_Fic_Mensal_VIS.txt_nficha
    Else
        TXT_NFICHA_CAD = frm_Alt_Fic_Mensal_VIS.txt_nficha
    End If
    TXT_CONTA_CAD = ""
    TXT_DESC_CAD = ""
    TXT_OP = ""
    TXT_DT_CAD = Date
    TXT_VALOR_CAD = 0
    
    GUIA.TabVisible(0) = True
    GUIA.TabVisible(1) = False
    
    de.rsTAB_DESC_CALC.Requery
    
    
'sql dos Crediarios em vencto
    w_mes = frm_Alt_Fic_Mensal_VIS.TXT_MES + 1
    w_ano = frm_Alt_Fic_Mensal_VIS.TXT_ANO
    W_DT = Format("01/" & w_mes & "/" & w_ano, "dd/mm/yyyy")
    W_DT = CVDate(W_DT) + 9


    'sql registros
    
    If de.rscmdBase.State = 1 Then de.rscmdBase.Close

 If acessoTotal() Then
        de.rscmdBase.Open "SELECT TAB_DESC_CALC.* FROM TAB_DESC_CALC WHERE (TAB_DESC_CALC.C_N_FICHA = " & TXT_NFICHA_CAD & ") ORDER BY TAB_DESC_CALC.C_TP_OP, C_DT", , adOpenStatic, adLockOptimistic
    Else
        de.rscmdBase.Open "SELECT TAB_DESC_CALC.* FROM TAB_DESC_CALC WHERE (TAB_DESC_CALC.C_N_FICHA = " & TXT_NFICHA_CAD & " AND TAB_DESC_CALC.C_TP_CONTA <> 20) ORDER BY TAB_DESC_CALC.C_TP_OP, C_DT", , adOpenStatic, adLockOptimistic
    End If

    Set adoReg.Recordset = de.rscmdBase.Clone
    de.rscmdBase.Close
    
    v_filtro = "VCTO <= #" & Format(CVDate(W_DT), "mm/dd/YYYY") & "#"
    v_filtro_puro = ""

        
    If Not adoReg.Recordset.EOF Then
 If acessoTotal() Then
        Set ADO_GRID.Recordset = de.cnc.Execute("SELECT Format(TAB_DESC_CALC.C_DT,'dd/mm/yy') AS DATA, TAB_DESC_CALC.C_TP_CONTA AS COD, TAB_TP_CONTA.TP_DESC AS TP, TAB_DESC_CALC.C_DESC AS CONTA, Format(TAB_DESC_CALC.C_VALOR,'R$ ###,##0.00') AS VALOR, TAB_DESC_CALC.C_TP_OP AS OP, Format(TAB_DESC_CALC.C_VISTO,'Yes/No') AS VISTO, TAB_DESC_CALC.C_CODIGO, TAB_DESC_CALC.C_N_FICHA, TAB_DESC_CALC.C_LOGO, TAB_DESC_CALC.C_NCRED, TAB_DESC_CALC.C_JUROS, TAB_DESC_CALC.C_VALOR_PARC, TAB_DESC_CALC.C_QTDE_PARC, TAB_DESC_CALC.C_EMP_COD, TAB_FICHA_MENS.M_NFICHA FROM TAB_TP_CONTA, TAB_DESC_CALC, TAB_FICHA_MENS WHERE (TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND TAB_DESC_CALC.C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.txt_nficha & " ) ORDER BY TAB_DESC_CALC.C_TP_OP, C_DT").Clone
    Else
        Set ADO_GRID.Recordset = de.cnc.Execute("SELECT Format(TAB_DESC_CALC.C_DT,'dd/mm/yy') AS DATA, TAB_DESC_CALC.C_TP_CONTA AS COD, TAB_TP_CONTA.TP_DESC AS TP, TAB_DESC_CALC.C_DESC AS CONTA, Format(TAB_DESC_CALC.C_VALOR,'R$ ###,##0.00') AS VALOR, TAB_DESC_CALC.C_TP_OP AS OP, Format(TAB_DESC_CALC.C_VISTO,'Yes/No') AS VISTO, TAB_DESC_CALC.C_CODIGO, TAB_DESC_CALC.C_N_FICHA, TAB_DESC_CALC.C_LOGO, TAB_DESC_CALC.C_NCRED, TAB_DESC_CALC.C_JUROS, TAB_DESC_CALC.C_VALOR_PARC, TAB_DESC_CALC.C_QTDE_PARC, TAB_DESC_CALC.C_EMP_COD, TAB_FICHA_MENS.M_NFICHA  FROM TAB_TP_CONTA, TAB_DESC_CALC, TAB_FICHA_MENS WHERE (TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND TAB_DESC_CALC.C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.txt_nficha & " AND TAB_DESC_CALC.C_TP_CONTA <> 20 ) ORDER BY TAB_DESC_CALC.C_TP_OP, C_DT").Clone
    End If
    End If
    V_MOVE = True

    Timer1.Enabled = True

   
    formatarFlexGrid
        

sair:
    Exit Sub
err1:
    If Not Err.Number = 3705 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

'*** Caption no navegador ***
Private Sub ADOREG_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1
    If Not adoReg.Recordset.EOF Then adoReg.Caption = "REGISTROS: " & adoReg.Recordset.AbsolutePosition & " / " & adoReg.Recordset.RecordCount & IIf(W_LD_FILTRO = True, " (FILTRADO)", "")
    
   If V_MOVE = True Then
        On Error Resume Next
        
       For I = 3 To 7
          
          If I > 3 And I <= 7 Then lbEmp(I).Visible = adoReg.Recordset.Fields("C_TP_CONTA") = 31
          If I >= 3 Then txt_Emp(I).Visible = adoReg.Recordset.Fields("C_TP_CONTA") = 31
       Next I
        
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
   
sair:
    V_MOVE = True
    Exit Sub
err1:
    If Not Err.Number = -2147217885 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub ADO_GRID_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1

    If V_MOVE = True Then

        V_MOVE = False
'        ADOREG.Recordset.Requery
        adoReg.Refresh
        adoReg.Recordset.Move ADO_GRID.Recordset.AbsolutePosition - 1

        V_MOVE = True
        

        
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
        Case "adicionar": Adicionar
        Case "editar":
            w_Editar = True
            Editar
                       If BarraF.Buttons("editar").Enabled = False Then txt_DT.SetFocus

        Case "salvar": Salvar
        Case "cancelar": Cancelar
        Case "excluir":   If BarraF.Buttons("adicionar").Enabled = True Then Excluir

'        Case "filtrar": FILTRAR
    End Select
End Sub


'*** Rotinas ***
Private Sub Adicionar()
On Error GoTo err1

    w_txt_desc = ""
    
    GUIA.TabEnabled(0) = False
    GUIA.TabVisible(0) = False
    GUIA.TabVisible(1) = True
    GUIA.TabEnabled(1) = True
    GUIA.Tab = 1
    
    w_PSS = w_PassWordLib

    Editar
    
    TXT_DT_CAD.SetFocus

    BarraF.Buttons("excluir").Enabled = False
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Cancelar()
On Error GoTo err1
    
If GUIA.TabVisible(0) = True Then   '*** alteração
    If adoReg.Recordset.RecordCount > 0 Then
    
        pos = adoReg.Recordset.AbsolutePosition - 1
        adoReg.Recordset.CancelBatch adAffectCurrent
        adoReg.Refresh
        adoReg.Recordset.Move pos
    
    End If
    Editar
    w_PSS = ""
    
    

    
Else '*** cad
    
    TXT_NFICHA_CAD = frm_Alt_Fic_Mensal_VIS.txt_nficha
    TXT_CONTA_CAD = ""
    TXT_DESC_CAD = ""
    TXT_OP = ""
    TXT_DT_CAD = Date
    TXT_VALOR_CAD = 0
        
    txt_Logo_Cad = ""
    txt_NCred_Cad = ""
    
    
    TXT_E_COD = ""
    TXT_E_JUROS = ""
    TXT_E_SALDO = ""
    TXT_E_VALOR = ""
    
    
    Editar
    w_PSS = ""

    GUIA.TabEnabled(0) = True
    GUIA.TabVisible(0) = True
    GUIA.TabEnabled(1) = False
    GUIA.TabVisible(1) = False
    GUIA.Tab = 0

    BarraF.Buttons("excluir").Enabled = True

End If

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub Editar()


If w_Editar Then
    On Error GoTo errPula
    
    
    If Not ADO_GRID.Recordset.State = 0 Then
    
        If ADO_GRID.Recordset.Fields("VISTO") = "Sim" Then
        frm_Habilitar.Show 1
        w_PSS = frm_Habilitar.txt_Pss
        Else
errPula:
        w_PSS = w_PassWordLib
        End If
    End If
End If
    
w_Editar = False

On Error GoTo err1

If w_PSS = w_PassWordLib Then
    

    w_At = False
    Pause 0.5
'SE FOR CREDIARIO MOSTRA O GRID DE CREDIARIOS
     If TXT_CONTA.BoundText = "17" And GUIA.TabVisible(0) = True Then
        MsgBox "Não é permitido editar Crediário, exclua e adicione novamente!", vbCritical
        GoTo sair
     ElseIf (TXT_CONTA.BoundText = "31" Or TXT_CONTA.BoundText = "9") And GUIA.TabVisible(0) = True Then
        MsgBox "Não é permitido editar Empréstimo, exclua e adicione novamente!", vbCritical
        GoTo sair
'     ElseIf TXT_CONTA.BoundText = "32" And GUIA.TabVisible(0) = True Then
 '       MsgBox "Não é permitido editar 13º, exclua e adicione novamente!", vbCritical
'        GoTo sair
     ElseIf TXT_CONTA.BoundText = "24" And GUIA.TabVisible(0) = True Then
        MsgBox "Não é permitido editar Férias, exclua e adicione novamente!", vbCritical
        GoTo sair
        
     Else
        GRID_CRED.Visible = False
        c_Filtro.Visible = False
     End If
    
    BarraF.Buttons("salvar").Enabled = Not BarraF.Buttons("salvar").Enabled
    BarraF.Buttons("cancelar").Enabled = Not BarraF.Buttons("cancelar").Enabled
    BarraF.Buttons("editar").Enabled = Not BarraF.Buttons("editar").Enabled
    BarraF.Buttons("adicionar").Enabled = Not BarraF.Buttons("adicionar").Enabled
    
    flexGRID.Enabled = Not flexGRID.Enabled
    
    GRID_CRED.Enabled = Not GRID_CRED.Enabled
    Grid.Enabled = Not Grid.Enabled
        
    GUIA.Enabled = Not GUIA.Enabled
   
    If BarraF.Buttons("salvar").Enabled = False Then flexGRID.SetFocus
    Pause 0.5
Else
    MsgBox "Senha de Liberação incorreta!", vbCritical
End If

sair:
    
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Excluir()
On Error GoTo err1
Dim w_exluir As Boolean
Dim fixo
        
w_At = False
        
'**Lançamento FIXO

w_excluir = False
fixo = Empty
On Error Resume Next
fixo = de.cnc.Execute("SELECT CF_EMP_COD FROM TAB_DESC_CALC_FIXO WHERE CF_EMP_COD = " & frm_Alt_Fic_Mensal_VIS.txt_F_COD & " AND CF_CODIGO = " & txt_ncred).Fields(0)
On Error GoTo err1
        
'If ADO_GRID.Recordset.Fields("C_VISTO") = True And w_PSS = "" Then
If ADO_GRID.Recordset.Fields("VISTO") = "Sim" Or fixo <> Empty Then
    frm_Habilitar.Show 1
    w_PSS = frm_Habilitar.txt_Pss
Else
    w_PSS = w_PassWordLib
End If

If w_PSS = w_PassWordLib Then
        

        
        If fixo <> Empty Then
            If vbNo = MsgBox("VOCÊ ESTÁ EXCLUINDO UM LANÇAMENTO PROGRAMADO, GERADO AUTOMATICAMENTE." & vbNewLine & "TEM CERTEZA QUE DESEJA EXCLUÍ-LO?", vbQuestion + vbYesNo, "Confirmação de Exclusão") Then
                Exit Sub
            End If
            w_excluir = True
        ElseIf vbYes = MsgBox("DESEJA REALMENTE EXCLUIR O LANÇAMENTO (" & TXT_CONTA & ")?", vbQuestion + vbYesNo) Then
            w_excluir = True
        End If
        
        If w_excluir = True Then
        
        '**FIM Lançamento Fixo
        
        'Se excluir o Emprestimo ,  tb  dá baixa no Cad Func.
      
            W_E_Cod = adoReg.Recordset.Fields("C_EMP_COD")
            W_C_COD = adoReg.Recordset.Fields("C_CODIGO")
            w_Valor = adoReg.Recordset.Fields("C_VALOR")
        
        
        If TXT_CONTA.BoundText = "9" And GUIA.TabVisible(0) = True Then
            '*** Qdo excluir Pag. de Emprestimo ! Pega o valor pago e desconta no saldo de emp. do tab_funcionario ***
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_EMPRESTIMO = F_EMPRESTIMO - '" & adoReg.Recordset.Fields("C_VALOR") + IIf(IsNull(adoReg.Recordset.Fields("c_JUROS")), 0, adoReg.Recordset.Fields("c_JUROS")) & "' WHERE (F_Codigo = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("m_F_COD") & ")"
            de.cnc.Execute "UPDATE TAB_EMPRESTIMO SET E_SALDO = E_SALDO - '" & adoReg.Recordset.Fields("C_VALOR") + adoReg.Recordset.Fields("c_JUROS") & "' WHERE E_Codigo = " & W_E_Cod & ""
        
        ElseIf TXT_CONTA.BoundText = "31" Then
            '*** Qdo excluir Emprestimo !   Pega o saldo q/ restou do emprestimo e desconta no saldo de emp. do tab_funcionario ***
            W_SALDO = 0
            On Error Resume Next
            W_SALDO = de.cnc.Execute("SELECT E_SALDO FROM TAB_EMPRESTIMO WHERE E_CODIGO = " & adoReg.Recordset.Fields("C_EMP_COD") & "").Fields("E_SALDO")
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_EMPRESTIMO =  F_EMPRESTIMO - '" & CDbl(W_SALDO) & "'  WHERE (F_Codigo = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("m_F_COD") & ")"
        End If
        
        
        
        If TXT_CONTA.BoundText = "9" Then
            '*** Excluir as Parcelas de Pagamentos de emprestimos
            
            
        ElseIf TXT_CONTA.BoundText = "31" Then
            
            '*** Excluir o Emprestimo
            de.cnc.Execute "DELETE * FROM TAB_EMPRESTIMO WHERE E_CODIGO = " & adoReg.Recordset.Fields("C_EMP_COD") & "", wRegAf
            If wRegAf = 0 Then MsgBox "Registro de Emprestimo não encontrado pra exclusão!", vbCritical
        
        End If
        
        w_Conta = TXT_CONTA.BoundText
     
        
        
        
        '*** Exclui o registro
        W_POS = adoReg.Recordset.AbsolutePosition - 1
        adoReg.Recordset.Delete
        w_adoFiltro = adoReg.Recordset.Filter
        Form_Load
        ADO_GRID.Refresh
        adoReg.Refresh
      
        adoReg.Recordset.Filter = w_adoFiltro
        ADO_GRID.Recordset.Filter = w_adoFiltro
      
        Grid.Refresh
      
        de.rsTAB_DESC_CALC.Close
        de.TAB_DESC_CALC
        '*** CALCULA O TOTAL - APÓS O NOVO VALOR ***
        
     If acessoTotal() Then
            W_MAIS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MAIS FROM TAB_DESC_CALC  WHERE (C_TP_OP = '+') AND (C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.txt_nficha & ")").Fields("MAIS")
            W_MENOS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MENOS FROM TAB_DESC_CALC WHERE (C_TP_OP = '-') AND (C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.txt_nficha & ")").Fields("MENOS")
        Else
            W_MAIS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MAIS FROM TAB_DESC_CALC  WHERE (C_TP_OP = '+') AND (C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.txt_nficha & ") AND (TAB_DESC_CALC.C_TP_CONTA <> 20)").Fields("MAIS")
            W_MENOS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MENOS FROM TAB_DESC_CALC WHERE (C_TP_OP = '-') AND (C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.txt_nficha & ") AND (TAB_DESC_CALC.C_TP_CONTA <> 20)").Fields("MENOS")
        End If
        
        W_TOTAL = IIf(IsNull(W_MAIS), 0, W_MAIS) + IIf(IsNull(W_MENOS), 0, W_MENOS)
        
        de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_TOTAL = '" & CDbl(W_TOTAL) & "' WHERE (M_NFICHA = " & frm_Alt_Fic_Mensal_VIS.txt_nficha & ")"
      
         w_PSS = ""
      '  Grid.ReBind
      
      
    '*** ATUALIZA A ULTIMA DATA DE PAGAMENTO NA TAB_EMPRESTIMO *** SE FOR DESCONTO
    If w_Conta = "9" Then
        W_DT_PG = de.cnc.Execute("SELECT MAX(EP_DT_PG) AS DT_ULT FROM TAB_EMPRESTIMO_PG WHERE (EP_Codigo = " & W_E_Cod & " AND EP_C_CODIGO <> " & frm_Alt_Fic_Mensal_VIS.txt_nficha & ")").Fields(0)
        If Not IsNull(W_DT_PG) Then
            de.cnc.Execute "UPDATE TAB_EMPRESTIMO SET E_DT_ULT_PG = '" & CVDate(W_DT_PG) & "', E_QT_PG = E_QT_PG-1  WHERE (E_Codigo = " & W_E_Cod & ")"
        Else
            de.cnc.Execute "UPDATE TAB_EMPRESTIMO SET E_DT_ULT_PG = E_DT_EMP, E_QT_PG = E_QT_PG-1  WHERE (E_Codigo = " & W_E_Cod & ")"
        End If
    ElseIf w_Conta = "14" Then
    
        '*** Atualiza VALOR DO SALDO DEVEDOR EM TAB_FUNCIONARIO ***
        de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_SALDO_ANT = F_SALDO_ANT + '" & CDbl(w_Valor) & "' WHERE (F_Codigo = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("m_F_COD") & ")"
    
    End If
        
        
        
        
        '*** ATUALIZAR A ANOTAÇÃO DO EMPRESTIMO DO FUNCIONARIO ***
        '** Sql EMP. P/ GRID
        If w_Conta = "9" Or w_Conta = "31" Then
            
            W_EMP_ANOT = ""
            Set ado_EMP.Recordset = de.cnc.Execute("SELECT * FROM TAB_EMPRESTIMO WHERE E_F_COD = " & frm_Alt_Fic_Mensal_VIS.TXT_FUNC.BoundText & " AND E_SALDO > 0").Clone
            Do While Not ado_EMP.Recordset.EOF
                W_EMP_ANOT = W_EMP_ANOT & IIf(Len(W_EMP_ANOT) > 0, vbCrLf, "") & ".  Dt Emp.: " & ado_EMP.Recordset.Fields("E_DT_EMP") & "           Juros : " & ado_EMP.Recordset.Fields("E_Juro_ao_mes") * 100 & " %           Dt Ult. Pg.: " & ado_EMP.Recordset.Fields("E_DT_ULT_PG")
                W_EMP_ANOT = W_EMP_ANOT & vbCrLf & ".  Parc. Pg.: " & ado_EMP.Recordset.Fields("E_QT_PG") & " / " & ado_EMP.Recordset.Fields("E_QT_PARC") & "       Valor Emp.: " & Format(ado_EMP.Recordset.Fields("E_VALOR"), "R$ 0.00") & "       SALDO: " & Format(ado_EMP.Recordset.Fields("E_SALDO"), "R$ 0.00")
            
                ado_EMP.Recordset.MoveNext
            Loop
            

            '*** UPDATE NO FUNCIONARIO ATUALIZANDO A ANOTAÇÃO DO EMPRESTIMO ***
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_EMPRESTIMO_ANOT = '" & W_EMP_ANOT & "' WHERE (F_Codigo = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("M_F_COD") & ")"
        End If
        
    End If
    formatarFlexGrid
    
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
        
       

        
        '*** CALCULA O TOTAL - APÓS O NOVO VALOR ***
     If acessoTotal() Then
            W_MAIS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MAIS FROM TAB_DESC_CALC  WHERE (C_TP_OP = '+') AND (C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.txt_nficha & ")").Fields("MAIS")
            W_MENOS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MENOS FROM TAB_DESC_CALC WHERE (C_TP_OP = '-') AND (C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.txt_nficha & ")").Fields("MENOS")
        Else
            W_MAIS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MAIS FROM TAB_DESC_CALC  WHERE (C_TP_OP = '+') AND (C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.txt_nficha & ") AND (TAB_DESC_CALC.C_TP_CONTA <> 20)").Fields("MAIS")
            W_MENOS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MENOS FROM TAB_DESC_CALC WHERE (C_TP_OP = '-') AND (C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.txt_nficha & ") AND (TAB_DESC_CALC.C_TP_CONTA <> 20)").Fields("MENOS")
        End If
        'W_MAIS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MAIS FROM TAB_DESC_CALC  WHERE (C_TP_OP = '+') AND (C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.txt_nficha & ")").Fields("MAIS")
        'W_MENOS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MENOS FROM TAB_DESC_CALC WHERE (C_TP_OP = '-') AND (C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.txt_nficha & ")").Fields("MENOS")
        
        W_TOTAL = IIf(IsNull(W_MENOS), 0, W_MENOS) + IIf(IsNull(W_MAIS), 0, W_MAIS)
        
        
        '***Pega saldo de emprestimo e atualiza ***
        w_Saldo_Emp = de.cnc.Execute("Select SUM(e_Saldo) as Saldo from Tab_Emprestimo where e_F_Cod = " & frm_Alt_Fic_Mensal_VIS.txt_F_COD & "").Fields(0)
        w_Saldo_Emp = IIf(IsNull(w_Saldo_Emp), 0, w_Saldo_Emp)
        
      
        '*** ATUALIZA NA TAB FICHA
        de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_TOTAL = '" & CDbl(W_TOTAL) & "' WHERE (M_NFICHA = " & frm_Alt_Fic_Mensal_VIS.txt_nficha & ")"
        '*** ATUALIZA SALDO DO EMPRESTIMO NO CAD. DE FUNC.
        de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_EMPRESTIMO = '" & CDbl(w_Saldo_Emp) & "' WHERE (F_CODIGO = " & frm_Alt_Fic_Mensal_VIS.txt_F_COD & ")"
       
        
        de.rsTAB_DESC_CALC.Requery
        de.rsTAB_DESC_CALC.Close
        de.TAB_DESC_CALC
        
        If de.rsTAB_FICHA_MENS.State = 1 Then de.rsTAB_FICHA_MENS.Requery
        frm_Alt_Fic_Mensal_VIS.Timer1 = True
        
sair:
    Unload Me
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub FILTRAR()
Dim w_resp As String
Dim W_CAMPO As String
Dim W_FILTRO As String
Dim W_FILTRO1 As String

On Error GoTo err1
    
    w_resp = InputBox("FILTRAR PELO QUÊ ? ENTRE COM O NÚMERO DA OPÇÃO DESEJADA." & Chr(13) & Chr(13) & "1 - Nº DA FICHA" & Chr(13) & "2 - FUNCIONÁRIO" & Chr(13) & "3 - MÊS E ANO" & Chr(13) & "4 - REMOVER FILTRO *", , "1")
    
    
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
Dim DB As dao.Database
Dim wtab As dao.Recordset
Dim wPARC As dao.Recordset


              
If GUIA.TabVisible(0) = True Then   '****   ALTERAR   ****
    
    'If Not (CDbl(Format(txt_DT, "mm")) = CDbl(frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("m_mes")) Or CDbl(Format(txt_DT, "mm")) = CDbl(frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("m_mes")) + 1) Then
        'MsgBox "Só é permitido data do mês passado ou do mês atual!", vbExclamation
    'Else
    
        If (txt_valor < 0 And (TXT_OP = "+" Or TXT_OP = "=")) Or (txt_valor > 0 And TXT_OP = "-") Then txt_valor = txt_valor * -1
        txt_valor.SetFocus
        txt_DT.SetFocus
        adoReg.Recordset.UpdateBatch adAffectCurrent
        
        If TXT_CONTA.BoundText = 32 Then
            'frm_Alt_Fic_Mensal_VIS.TXT_13_OBS = TXT_OBS
            'frm_Alt_Fic_Mensal_VIS.TXT_13_PG = txt_13
            '*** Atualiza Dt 13º ***  TAB_FUNCIONARIO
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_13_OK = 0 , F_13_PG = '" & txt_13 & "' , F_13_OBS = '" & TXT_OBS & "' WHERE (F_Codigo = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("M_F_COD") & ")"
            '*** Atualiza Dt 13º***   TAB_FICHA_MENS
            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_13_OK = 0 , M_13_PG = '" & txt_13 & "', M_13_OBS = '" & TXT_OBS & "'  WHERE (M_F_Cod = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("M_F_COD") & ")"
        End If
        
       'Atualiza no cadastro dos fixos
       If flexGRID.TextMatrix(flexGRID.RowSel, 10) > 0 Then
            'cf_codigo, cf_emp_cod, cf_dt, cf_valor, cf_tp_conta, cf_tp_op, cf_desc, cf_visto
            de.cnc.Execute ("UPDATE TAB_DESC_CALC_FIXO SET  cf_dt = '" & txt_DT & "', cf_valor = '" & CDbl(txt_valor) & "', cf_tp_conta = " & TXT_CONTA_cod & ", cf_tp_op = '" & TXT_OP & "', cf_desc = '" & TXT_DESC & "' WHERE cf_codigo = " & flexGRID.TextMatrix(flexGRID.RowSel, 10))

       End If
        
        
        
        Editar
        w_At = True
            
        'Pause 1
        adoReg.Refresh
        
        de.rsTAB_DESC_CALC.Requery
        If de.rscmdBase.State = 1 Then de.rscmdBase.Close
    
     If acessoTotal() Then
            de.rscmdBase.Open "SELECT TAB_DESC_CALC.* FROM TAB_DESC_CALC WHERE (TAB_DESC_CALC.C_N_FICHA = " & TXT_NFICHA_CAD & ") ORDER BY TAB_DESC_CALC.C_TP_OP, C_DT", , adOpenStatic, adLockOptimistic
        Else
            de.rscmdBase.Open "SELECT TAB_DESC_CALC.* FROM TAB_DESC_CALC WHERE (TAB_DESC_CALC.C_N_FICHA = " & TXT_NFICHA_CAD & ") AND (TAB_DESC_CALC.C_TP_CONTA <> 20) ORDER BY TAB_DESC_CALC.C_TP_OP, C_DT", , adOpenStatic, adLockOptimistic
        End If
    
    
    
        Set adoReg.Recordset = de.rscmdBase.Clone
        
        adoReg.Refresh
        
     If acessoTotal() Then
            Set ADO_GRID.Recordset = de.cnc.Execute("SELECT Format(TAB_DESC_CALC.C_DT,'dd/mm/yy') AS DATA, TAB_DESC_CALC.C_TP_CONTA AS COD, TAB_TP_CONTA.TP_DESC AS TP, TAB_DESC_CALC.C_DESC AS CONTA, Format(TAB_DESC_CALC.C_VALOR,'R$ ###,##0.00') AS VALOR, TAB_DESC_CALC.C_TP_OP AS OP, Format(TAB_DESC_CALC.C_VISTO,'Yes/No') AS VISTO, TAB_DESC_CALC.C_CODIGO, TAB_DESC_CALC.C_N_FICHA, TAB_DESC_CALC.C_LOGO, TAB_DESC_CALC.C_NCRED, TAB_DESC_CALC.C_JUROS, TAB_DESC_CALC.C_VALOR_PARC, TAB_DESC_CALC.C_QTDE_PARC, TAB_DESC_CALC.C_EMP_COD, TAB_FICHA_MENS.M_NFICHA FROM TAB_TP_CONTA, TAB_DESC_CALC, TAB_FICHA_MENS WHERE (TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND TAB_DESC_CALC.C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.txt_nficha & " ) ORDER BY TAB_DESC_CALC.C_TP_OP, C_DT").Clone
        Else
            Set ADO_GRID.Recordset = de.cnc.Execute("SELECT Format(TAB_DESC_CALC.C_DT,'dd/mm/yy') AS DATA, TAB_DESC_CALC.C_TP_CONTA AS COD, TAB_TP_CONTA.TP_DESC AS TP, TAB_DESC_CALC.C_DESC AS CONTA, Format(TAB_DESC_CALC.C_VALOR,'R$ ###,##0.00') AS VALOR, TAB_DESC_CALC.C_TP_OP AS OP, Format(TAB_DESC_CALC.C_VISTO,'Yes/No') AS VISTO, TAB_DESC_CALC.C_CODIGO, TAB_DESC_CALC.C_N_FICHA, TAB_DESC_CALC.C_LOGO, TAB_DESC_CALC.C_NCRED, TAB_DESC_CALC.C_JUROS, TAB_DESC_CALC.C_VALOR_PARC, TAB_DESC_CALC.C_QTDE_PARC, TAB_DESC_CALC.C_EMP_COD, TAB_FICHA_MENS.M_NFICHA  FROM TAB_TP_CONTA, TAB_DESC_CALC, TAB_FICHA_MENS WHERE (TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND TAB_DESC_CALC.C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.txt_nficha & " AND TAB_DESC_CALC.C_TP_CONTA <> 20 ) ORDER BY TAB_DESC_CALC.C_TP_OP, C_DT").Clone
        End If

        ADO_GRID.Refresh
        'Pause 1
        Grid.ReBind
    
        W_FICHA = txt_nficha
        w_PSS = ""
        
        
    'End If
    
    
    
Else    '**** CADASTRAR ****

    
    If lb_form = "mensal" Then
        w_mes = frm_Alt_Fic_Mensal_VIS.TXT_MES
        w_ano = frm_Alt_Fic_Mensal_VIS.TXT_ANO
    ElseIf lb_form = "visualizar" Then
        w_mes = frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("M_MES")
        w_ano = frm_Alt_Fic_Mensal_VIS.TXT_ANO
    End If
    
    If TXT_DT_EXTRA.Visible = True And TXT_DT_EXTRA = "" Then
        MsgBox "Você deve preencher Data do pagamento!  " & LB_DT_EXTRA, vbInformation
        TXT_DT_EXTRA.SetFocus
        GoTo sair
    End If
    
    If (CDbl(Format(TXT_DT_CAD, "mm")) >= CDbl(w_mes) Or CDbl(Format(TXT_DT_CAD, "mm")) = CDbl(w_mes) - 1) And (Not CDbl(TXT_VALOR_CAD) = 0 Or TXT_VALOR_CAD <> "") And TXT_OP_CAD <> "" Then
        
        
        If (TXT_VALOR_CAD < 0 And TXT_OP_CAD = "+") Or (TXT_VALOR_CAD > 0 And TXT_OP_CAD = "-") Then TXT_VALOR_CAD = TXT_VALOR_CAD * -1
        
        If txt_Logo_Cad = "" And TXT_CONTA_CAD.BoundText <> "31" Then   '*** NÃO SEJA CREDIARIOS E NÃO SEJA EMPRESTIMO ***
                de.cmdIncluirDescCalc TXT_DT_CAD, TXT_NFICHA_CAD, TXT_CONTA_CAD.BoundText, TXT_OP_CAD, TXT_VALOR_CAD, IIf(TXT_DESC_CAD = "", " ", TXT_DESC_CAD), "0", "0", IIf(TXT_CONTA_CAD.BoundText = "9", CDbl(IIf(TXT_E_JUROS = "", 0, TXT_E_JUROS)), "0"), "0", CDbl(IIf(TXT_E_COD = "", 0, TXT_E_COD))
        ElseIf TXT_CONTA_CAD.BoundText <> "31" Then  '*** DIFERENTE DE EMPRESTIMO ***
            de.cmdIncluirDescCalc TXT_DT_CAD, TXT_NFICHA_CAD, TXT_CONTA_CAD.BoundText, TXT_OP_CAD, TXT_VALOR_CAD, TXT_DESC_CAD, txt_Logo_Cad, txt_NCred_Cad, "0", "0", CDbl(IIf(TXT_E_COD = "", 0, TXT_E_COD))
        
        ElseIf TXT_CONTA_CAD.BoundText = "31" Then  'SEJA EMPRESTIMO ***
            
            If txt_Emp(0) <> "" And txt_Emp(1) <> "" Then
                
                
        '       *** SALVA O EMPRESTIMO ***
                W_DT = CVDate(TXT_DT_CAD)
                'de.cmdIncluirEmprestimo frm_Alt_Fic_Mensal_VIS.TXT_FUNC.BoundText, W_DT, CDbl(txt_Emp(1)) / 100, (CDbl(txt_Emp(1)) / 100) / 30, txt_Emp(0), CDbl(TXT_VALOR_CAD), CDbl(TXT_VALOR_CAD), 0, W_DT, txt_Emp(2)
                de.cmdIncluirEmprestimo frm_Alt_Fic_Mensal_VIS.TXT_FUNC.BoundText, W_DT, CDbl(txt_Emp(1)), CDbl(txt_Emp(1)) / 30, txt_Emp(0), CDbl(TXT_VALOR_CAD), CDbl(TXT_VALOR_CAD), 0, W_DT, txt_Emp(2)
                'PEGA O CODIGO DO EMPRESTIMO
                W_E_Cod = de.cnc.Execute("Select MAX(E_CODIGO) as COD from TAB_EMPRESTIMO").Fields(0)
                
                
                '*** SALVA A CONTA ***
                de.cmdIncluirDescCalc TXT_DT_CAD, TXT_NFICHA_CAD, TXT_CONTA_CAD.BoundText, TXT_OP_CAD, CDbl(TXT_VALOR_CAD), TXT_DESC_CAD, "0", "0", CDbl(txt_Emp(1)), txt_Emp(0), W_E_Cod
            Else
                MsgBox "Preencha a Qtde de Parcelas e Juros !", vbExclamation
                txt_Emp(0).SetFocus
                Exit Sub
            End If
            
        
        End If
        
        Select Case TXT_CONTA_CAD.BoundText
        '*** atualiza Data de PG de Férias
        Case 24:

            '*** Atualiza Dt 13º ***  TAB_FUNCIONARIO
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS_OK = 0 , F_FERIAS_ULT_PG = F_FERIAS_PG, F_FERIAS_PG = '" & TXT_DT_EXTRA & "', F_FERIAS = '" & TXT_DESC_EXTRA & "'  WHERE (F_Codigo = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("M_F_COD") & ")"
            '*** Atualiza Dt 13º***   TAB_FICHA_MENS
            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_FERIAS_OK = 0 , M_FERIAS_ULT_PG = M_FERIAS_PG, M_FERIAS_PG = '" & TXT_DT_EXTRA & "',M_FERIAS = '" & TXT_DESC_EXTRA & "'  WHERE (M_NFICHA = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("M_NFICHA") & ")"
        
        
        '*** atualiza Data de PG de 13º
        Case 32:
            '*** Atualiza Dt 13º ***  TAB_FUNCIONARIO
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_13_OK = 0 , F_13_ULT_PG = F_13_PG, F_13_PG = '" & TXT_DT_EXTRA & "' , F_13_OBS = '" & TXT_DESC_EXTRA & "' WHERE (F_Codigo = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("M_F_COD") & ")"
            '*** Atualiza Dt 13º***   TAB_FICHA_MENS
            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_13_OK = 0 , M_13_ULT_PG = M_13_PG, M_13_PG = '" & TXT_DT_EXTRA & "', M_13_OBS = '" & TXT_DESC_EXTRA & "'  WHERE (M_F_Cod = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("M_F_COD") & ")"
                
        
        '*** Desconto Saldo Mês Anterior
        Case 14:
            
            '*** Atualiza VALOR DO SALDO DEVEDOR EM TAB_FUNCIONARIO ***
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_SALDO_ANT = F_SALDO_ANT - '" & TXT_VALOR_CAD & "' WHERE (F_Codigo = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("M_F_COD") & ")"
                        
        
        '*** Desconto Emprestimo   OU   EMPRESTIMO
        Case 9:
            '*** Cadastra as Parcela paga na Tabela de Pagamento de Emprestimos ***
            w_parc = de.cnc.Execute("Select Count(ep_codigo) as Qtde from Tab_Emprestimo_PG Where EP_CODIGO = " & TXT_E_COD & " and EP_PARC <> 0").Fields(0)
            w_qt_dias = CDbl(CVDate(TXT_DT_CAD) - CVDate(ado_EMP.Recordset.Fields("E_DT_ULT_PG")))
            w_Valor = (CDbl(TXT_VALOR_CAD) + CDbl(TXT_E_JUROS)) * -1
            If (CDbl(TXT_VALOR_CAD) * -1) > CDbl(TXT_E_JUROS) Then w_parc = w_parc + 1
            
            '*** Dá baixa no emprestimo na tab. funcionario ***
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_EMPRESTIMO = F_EMPRESTIMO - '" & CDbl(w_Valor) & "' WHERE (F_Codigo = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("M_F_COD") & ")"
            
            '*** Dá baixa no emprestimo na tab. emprestimo ***
            de.cnc.Execute "UPDATE TAB_EMPRESTIMO SET E_QT_PG = E_QT_PG + 1 , E_DT_ULT_PG = '" & TXT_DT_CAD & "', E_SALDO = E_SALDO - '" & CDbl(w_Valor) & "' WHERE (E_Codigo = " & TXT_E_COD & ")"
            
            W_C_CODIGO = de.cnc.Execute("SELECT MAX(C_CODIGO) AS C_COD FROM TAB_DESC_CALC").Fields(0)
            
            '*** Inclui conta na Ficha
            de.cmdIncluirEmprestimoPG TXT_E_COD, TXT_DT_CAD, w_parc, w_qt_dias, CDbl(w_Valor), CDbl(TXT_E_JUROS), W_C_CODIGO
            
        '*** EMPRESTIMO
        Case 31:
            '*** Dá Entrada(soma) no emprestimo na tab. funcionario ***
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_EMPRESTIMO = F_EMPRESTIMO + '" & CDbl(TXT_VALOR_CAD) & "' WHERE (F_Codigo = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("M_F_COD") & ")"
        
        End Select
    
     
    '*** ATUALIZAR A ANOTAÇÃO DO EMPRESTIMO DO FUNCIONARIO ***
        '** Sql EMP. P/ GRID
        If TXT_CONTA_Cod_CAD.BoundText = "9" Or TXT_CONTA_Cod_CAD.BoundText = "31" Then
            TXT_E_JUROS = IIf(TXT_E_JUROS = "", 0, TXT_E_JUROS)
            
            W_EMP_ANOT = ""
            Set ado_EMP.Recordset = de.cnc.Execute("SELECT * FROM TAB_EMPRESTIMO WHERE E_F_COD = " & frm_Alt_Fic_Mensal_VIS.TXT_FUNC.BoundText & " AND (E_SALDO > 0  OR E_DT_ULT_PG <= #" & Format(TXT_DT_CAD, "MM/DD/YYYY") & "#)").Clone
            Do While Not ado_EMP.Recordset.EOF
           
                W_EMP_ANOT = W_EMP_ANOT & IIf(Len(W_EMP_ANOT) > 0, vbCrLf, "") & ". Dt Emp.: " & ado_EMP.Recordset.Fields("E_DT_EMP") & "    Valor Emp.: " & Format(ado_EMP.Recordset.Fields("E_VALOR"), "R$ 0.00") & "     Juros : " & ado_EMP.Recordset.Fields("E_Juro_ao_mes") * 100 & " %" & "     Parc. Pg.: " & ado_EMP.Recordset.Fields("E_QT_PG") & " / " & ado_EMP.Recordset.Fields("E_QT_PARC")
                W_EMP_ANOT = W_EMP_ANOT & vbCrLf & ". Saldo Ant.: " & Format(CDbl(ado_EMP.Recordset.Fields("E_SALDO")) - IIf(TXT_CONTA_Cod_CAD.BoundText = "9", CDbl(TXT_VALOR_CAD) + CDbl(TXT_E_JUROS), 0), "R$ 0.00") & "         Dt Ult. Pg.: " & ado_EMP.Recordset.Fields("E_DT_ULT_PG") & "        Saldo At.: " & Format(CDbl(ado_EMP.Recordset.Fields("E_SALDO")), "R$ 0.00")
            
            
                ado_EMP.Recordset.MoveNext
            Loop
            
            
            '*** UPDATE NO FUNCIONARIO ATUALIZANDO A ANOTAÇÃO DO EMPRESTIMO ***
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_EMPRESTIMO_ANOT = '" & W_EMP_ANOT & "' WHERE (F_Codigo = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("M_F_COD") & ")"
            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_EMPRESTIMO_ANOT = '" & W_EMP_ANOT & "' WHERE (M_NFICHA = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("M_NFICHA") & ")"
        End If
     
     
     
     
     
     
     
        '*** salvar nos lojbs ***
'        If de.rsLOJB118.State = 0 Then de.LOJB118
'        de.cncPDX.Execute "INSERT INTO LOJB118(CRED_LOJA, N_CRED, PARCELA, CONTROLE, VALOR, JUROS, DATA_PGT, COD_LOJ, DATA_AT) VALUES ('" & txt_Logo_Cad & "'," & txt_NCred_Cad & ", '1', '1', '" & Format(TXT_VALOR_CAD, "0.00") & "', 0, '" & Format(TXT_DT_CAD, "dd/mm/yyyy") & "', '" & txt_Logo_Cad & "', '" & Format(Date, "dd/mm/yyyy") & "')", RegAf
                
        MsgBox "Registro salvo com sucesso!", vbInformation
        
        TXT_CONTA_CAD = ""
        
        BarraF.Buttons("excluir").Enabled = True
        Editar
        Form_Load
        
        W_FICHA = TXT_NFICHA_CAD
        
        
        
        
        
        
        
    ElseIf Not (CDbl(Format(TXT_DT_CAD, "mm")) = CDbl(frm_Alt_Fic_Mensal_VIS.TXT_MES) Or CDbl(Format(TXT_DT_CAD, "mm")) = CDbl(frm_Alt_Fic_Mensal_VIS.TXT_MES) - 1) Then
        MsgBox "Só é permitido data do mês passado ou do mês atual!", vbExclamation
    Else
        MsgBox "Preencha os Campos!", vbCritical
    End If


End If

   If txt_nficha <> "" Then
        
        de.rsTAB_DESC_CALC.Close
        de.TAB_DESC_CALC
        
        '*** CALCULA O TOTAL - APÓS O NOVO VALOR ***
        W_MAIS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MAIS FROM TAB_DESC_CALC  WHERE (C_TP_OP = '+') AND (C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.txt_nficha & ")").Fields("MAIS")
        W_MENOS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MENOS FROM TAB_DESC_CALC WHERE (C_TP_OP = '-') AND (C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.txt_nficha & ")").Fields("MENOS")
        
        W_TOTAL = IIf(IsNull(W_MAIS), 0, W_MAIS) - IIf(IsNull(W_MENOS), 0, W_MENOS)

            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_TOTAL = '" & CDbl(W_TOTAL) & "' WHERE (M_NFICHA = " & txt_nficha & ")"
  End If


formatarFlexGrid

sair:

    Exit Sub
err1:
    If Err.Number = -2147467259 Then
        MsgBox "Este item já foi incluído na ficha!", vbExclamation
    Else
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    End If
    
    Resume sair
       
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Fechar
End Sub

Private Sub GRID_CRED_DblClick()
    mnuSelSel_Click
End Sub

Private Sub GRID_CRED_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo err1
    
    If Button = 2 Then PopupMenu mnuSel


sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub











Private Sub GRID_EMP_DblClick()
On Error GoTo err1

If Not ado_EMP.Recordset.EOF Then

    c_Filtro_Emp.value = 0
    
    Dim WADOEMP As ADODB.Recordset
    
    Set WADOEMP = ado_EMP.Recordset.Clone
    WADOEMP.Move ado_EMP.Recordset.AbsolutePosition - 1
    
    TXT_E_COD = WADOEMP.Fields("E_Codigo")
    TXT_E_SALDO = Format(WADOEMP.Fields("E_Saldo"), "R$ 0.00")
    TXT_E_VALOR = Format(WADOEMP.Fields("E_Valor"), "R$ 0.00")
    
        
    
    TXT_E_JUROS = Format(CALC_PG_EMP(WADOEMP, TXT_DT_CAD), "R$ 0.00")
    W_PARC_RESTANTE = WADOEMP.Fields("E_QT_PARC") - WADOEMP.Fields("E_QT_PG")
    TXT_VALOR_CAD = (CDbl(TXT_E_SALDO) / W_PARC_RESTANTE) + CDbl(TXT_E_JUROS)
    
    w_txt_desc = "Pg. Emp.: " & WADOEMP.Fields("E_QT_PG") + 1 & "/" & WADOEMP.Fields("E_QT_PARC")
    TXT_DESC_CAD = "Pg. Emp.: " & WADOEMP.Fields("E_QT_PG") + 1 & "/" & WADOEMP.Fields("E_QT_PARC") & vbCrLf & "Valor : " & Format(TXT_VALOR_CAD - TXT_E_JUROS, "R$ 0.00") & "    +    Juros : " & Format(TXT_E_JUROS, "R$ 0.00")

End If
    
sair:

    Exit Sub
err1:
    Resume sair
    
End Sub
















Private Sub GUIA_GotFocus()
    If GUIA.Tab = 0 Then
        txt_DT.SetFocus
    Else
        TXT_DT_CAD.SetFocus
    End If
End Sub



Private Sub mnuSelSel_Click()
On Error GoTo err1

    If GUIA.Tab = 0 Then  'Alteração
    
        TXT_LOGO = ADO_CRED.Recordset.Fields("cred_loja")
        TXT_NUM = ADO_CRED.Recordset.Fields("n_cred")
        txt_valor = Format(ADO_CRED.Recordset.Fields("SALDO"), "R$ 0.00")
        TXT_DESC = "CT. : " & TXT_LOGO & "." & TXT_NUM & "   -   DATA VCTO : " & ADO_CRED.Recordset.Fields("VCTO") & vbCrLf & "    -    VALOR : " & Format(txt_valor, "R$ 0.00") & "    -    SALDO : " & Format(ADO_CRED.Recordset.Fields("SALDO"), "R$ 0.00")
    Else    'Cadastro
    
        txt_Logo_Cad = Mid(ADO_CRED.Recordset.Fields("cred"), 1, 2)
        txt_NCred_Cad = Mid(ADO_CRED.Recordset.Fields("cred"), 4)
        TXT_VALOR_CAD = Format(ADO_CRED.Recordset.Fields("SALDO"), "R$ 0.00")
        TXT_DESC_CAD = "CT. : " & txt_Logo_Cad & "." & txt_NCred_Cad & "   -   DT. VCTO : " & ADO_CRED.Recordset.Fields("VCTO") & vbCrLf & "VALOR : " & Format(TXT_VALOR_CAD, "R$ 0.00") & "    -    SALDO : " & Format(ADO_CRED.Recordset.Fields("SALDO"), "R$ 0.00")
        
    End If

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair

End Sub

Private Sub Timer1_Timer()
On Error Resume Next

    If adoReg.Recordset.State = 1 Then If adoReg.Recordset.EOF Then Adicionar
    Timer1.Enabled = False

End Sub



Private Sub TXT_13_OBS_KeyDown(KeyCode As Integer, Shift As Integer)
       If KeyCode = 13 Then If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1) Then Salvar
        
End Sub

Private Sub TXT_CONTA_CAD_Change()
On Error GoTo err1
        
   If w_At = False Then
       w_At = True
        TXT_CONTA_CAD_op.BoundText = TXT_CONTA_CAD.BoundText
        TXT_CONTA_Cod_CAD.BoundText = TXT_CONTA_CAD.BoundText
        TXT_OP_CAD = TXT_CONTA_CAD_op.Text
       w_At = False
   End If

        '** Sql EMP. P/ GRID
        If TXT_CONTA_Cod_CAD.BoundText = "9" Then Set ado_EMP.Recordset = de.cnc.Execute("SELECT * FROM TAB_EMPRESTIMO WHERE E_F_COD = " & frm_Alt_Fic_Mensal_VIS.TXT_FUNC.BoundText & " AND E_SALDO > 0").Clone
            
         GRID_EMP.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
    
         LB_EMP_D(0).Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         LB_EMP_D(1).Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         LB_EMP_D(2).Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         LB_EMP_D(3).Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         TXT_E_COD.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         TXT_E_SALDO.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         TXT_E_VALOR.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         TXT_E_JUROS.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         c_Filtro_Emp.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"

         If TXT_CONTA_Cod_CAD.BoundText = "24" Then
            LB_DT_EXTRA = "DT. (F)"
            LB_DESC_EXTRA = "DESCRIÇÃO DAS FÉRIAS"
         ElseIf TXT_CONTA_Cod_CAD.BoundText = "32" Then
            LB_DT_EXTRA = "DT. (13º)"
            LB_DESC_EXTRA = "OBS 13º"
         End If
         
         LB_DT_EXTRA.Visible = TXT_CONTA_Cod_CAD.BoundText = "24" Or TXT_CONTA_Cod_CAD.BoundText = "32"
         TXT_DT_EXTRA.Visible = TXT_CONTA_Cod_CAD.BoundText = "24" Or TXT_CONTA_Cod_CAD.BoundText = "32"
         LB_DESC_EXTRA.Visible = TXT_CONTA_Cod_CAD.BoundText = "24" Or TXT_CONTA_Cod_CAD.BoundText = "32"
         TXT_DESC_EXTRA.Visible = TXT_CONTA_Cod_CAD.BoundText = "24" Or TXT_CONTA_Cod_CAD.BoundText = "32"
       
    
    If TXT_CONTA_CAD.BoundText = "31" Then '*** EMPRESTAR ***
        'Emprestimo
        
        
    ElseIf TXT_CONTA_CAD.BoundText = "9" Then  '*** DESCONTO DO EMPRESTIMO ***
        'Emprestimo descontos
        
    End If
    
    If UCase(TXT_CONTA_CAD.Text) = "DESC. CRED." Then
        
        CREDIARIO
        
        txt_NCred_Cad.Visible = True
        txt_Logo_Cad.Visible = True
        lbncred_cad.Visible = True
        lblogo_cad.Visible = True
        c_Filtro.Visible = True
        GRID_CRED.Visible = True
        
    Else
        txt_NCred_Cad.Visible = False
        txt_Logo_Cad.Visible = False
        lbncred_cad.Visible = False
        lblogo_cad.Visible = False
        GRID_CRED.Visible = False
        c_Filtro.Visible = False
    End If
    
sair:
    Exit Sub
err1:
    If Err.Number = -2147467259 Then
        MsgBox "Este item já foi incluído na ficha!", vbExclamation
    Else
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    End If
    Resume sair
End Sub



Private Sub TXT_CONTA_CAD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub TXT_CONTA_CAD_Validate(Cancel As Boolean)
           
           
        lbEmp(0).Visible = TXT_CONTA_CAD.BoundText = "31"
        lbEmp(1).Visible = TXT_CONTA_CAD.BoundText = "31"
        lbEmp(2).Visible = TXT_CONTA_CAD.BoundText = "31"
        lbEmp(3).Visible = TXT_CONTA_CAD.BoundText = "31"
        
        txt_Emp(0).Visible = TXT_CONTA_CAD.BoundText = "31"
        txt_Emp(1).Visible = TXT_CONTA_CAD.BoundText = "31"
        txt_Emp(2).Visible = TXT_CONTA_CAD.BoundText = "31"
        txt_Emp(2) = Format(TXT_DT_CAD, "DD")
        
End Sub

Private Sub TXT_CONTA_Change()
On Error GoTo err1
   If w_At = False And BarraF.Buttons("cancelar").Enabled = True Then
            
           If adoReg.Recordset.Fields("C_TP_CONTA") <> TXT_CONTA.BoundText Then
                txt_conta_Op.BoundText = TXT_CONTA.BoundText
                
                TXT_CONTA_cod.BoundText = TXT_CONTA.BoundText
                TXT_OP = txt_conta_Op.Text
           End If
   End If
    
  If TXT_CONTA.BoundText <> "" Then
        If TXT_CONTA.BoundText = 17 Then
            TXT_NUM.Visible = True
            TXT_LOGO.Visible = True
            LBNCRED.Visible = True
            LBLOGO.Visible = True
            
            If BarraF.Buttons("editar").Enabled = False Then
                GRID_CRED.Visible = True
                c_Filtro.Visible = True
                
                MsgBox "O sistema não permite alterar , um item p/ crediário! Somente incluir!", vbExclamation
                Cancelar
            End If
        
        ElseIf TXT_CONTA.BoundText = 31 Then
    
            If BarraF.Buttons("editar").Enabled = False Then
                MsgBox "O sistema não permite alterar, um item p/ Emprestimo!", vbExclamation
                Cancelar
            End If
        Else
            TXT_NUM.Visible = False
            TXT_LOGO.Visible = False
            LBNCRED.Visible = False
            LBLOGO.Visible = False
            c_Filtro.Visible = False
            GRID_CRED.Visible = False
        End If
    End If

sair:
    Exit Sub
err1:
    If Err.Number = -2147467259 Then
        MsgBox "Este item já foi incluído na ficha!", vbExclamation
    Else
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    End If
    Resume sair
    
End Sub


'--------- Ao Pressionar uma Tecla -----------
Private Sub GUIA_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub TXT_CONTA_Cod_CAD_Change()
   If w_At = False Then
       w_At = True
       TXT_CONTA_CAD.BoundText = TXT_CONTA_Cod_CAD.BoundText
       TXT_CONTA_CAD_op.BoundText = TXT_CONTA_CAD.BoundText
       TXT_OP_CAD = TXT_CONTA_CAD_op
                       
        '** Sql EMP. P/ GRID
        If TXT_CONTA_Cod_CAD.BoundText = "9" Then Set ado_EMP.Recordset = de.cnc.Execute("SELECT * FROM TAB_EMPRESTIMO WHERE E_F_COD = " & frm_Alt_Fic_Mensal_VIS.TXT_FUNC.BoundText & " AND E_SALDO > 0").Clone
            
         GRID_EMP.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
    
         LB_EMP_D(0).Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         LB_EMP_D(1).Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         LB_EMP_D(2).Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         LB_EMP_D(3).Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         TXT_E_COD.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         TXT_E_SALDO.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         TXT_E_VALOR.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         TXT_E_JUROS.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         c_Filtro_Emp.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         
         
         If TXT_CONTA_Cod_CAD.BoundText = "24" Then
            LB_DT_EXTRA = "DT. (F)"
            LB_DESC_EXTRA = "DESCRIÇÃO DAS FÉRIAS"
         ElseIf TXT_CONTA_Cod_CAD.BoundText = "32" Then
            LB_DT_EXTRA = "DT. (13º)"
            LB_DESC_EXTRA = "OBS 13º"
         End If
         
         LB_DT_EXTRA.Visible = TXT_CONTA_Cod_CAD.BoundText = "24" Or TXT_CONTA_Cod_CAD.BoundText = "32"
         TXT_DT_EXTRA.Visible = TXT_CONTA_Cod_CAD.BoundText = "24" Or TXT_CONTA_Cod_CAD.BoundText = "32"
         LB_DESC_EXTRA.Visible = TXT_CONTA_Cod_CAD.BoundText = "24" Or TXT_CONTA_Cod_CAD.BoundText = "32"
         TXT_DESC_EXTRA.Visible = TXT_CONTA_Cod_CAD.BoundText = "24" Or TXT_CONTA_Cod_CAD.BoundText = "32"
       
       w_At = False
   End If
   
End Sub


Private Sub TXT_CONTA_Cod_CAD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub TXT_CONTA_Cod_CAD_Validate(Cancel As Boolean)
           
        lbEmp(0).Visible = TXT_CONTA_CAD.BoundText = "31"
        lbEmp(1).Visible = TXT_CONTA_CAD.BoundText = "31"
        lbEmp(2).Visible = TXT_CONTA_CAD.BoundText = "31"
        lbEmp(3).Visible = TXT_CONTA_CAD.BoundText = "31"
        
        txt_Emp(0).Visible = TXT_CONTA_CAD.BoundText = "31"
        txt_Emp(1).Visible = TXT_CONTA_CAD.BoundText = "31"
        txt_Emp(2).Visible = TXT_CONTA_CAD.BoundText = "31"
        
        
        txt_Emp(2) = Format(TXT_DT_CAD, "DD")
           
End Sub

Private Sub TXT_CONTA_COD_Change()
   
   If w_At = False Then
       w_At = True
        TXT_CONTA.BoundText = TXT_CONTA_cod.BoundText
        txt_conta_Op.BoundText = TXT_CONTA.BoundText
        TXT_OP = txt_conta_Op
       w_At = False
    End If
    
    TXT_E_COD_E.Visible = (TXT_CONTA.BoundText = "9" Or TXT_CONTA.BoundText = "31")
    LB_EMP_DE(0).Visible = (TXT_CONTA.BoundText = "9" Or TXT_CONTA.BoundText = "31")
    TXT_E_JUROS_E.Visible = (TXT_CONTA.BoundText = "9")
    LB_EMP_DE(1).Visible = (TXT_CONTA.BoundText = "9")
    
    If TXT_CONTA_cod.BoundText = "24" Then
       lb_dt_13 = "DT. (F)"
       lb_OBS = "DESCRIÇÃO DAS FÉRIAS"
       lb_dt_13.Visible = True
       lb_OBS.Visible = True
       txt_13.Visible = True
       TXT_OBS.Visible = True
    ElseIf TXT_CONTA_cod.BoundText = "32" Then
       lb_dt_13 = "DT. (13º)"
       lb_OBS = "OBS 13º"
       lb_dt_13.Visible = True
       lb_OBS.Visible = True
       txt_13.Visible = True
       txt_13 = frm_Alt_Fic_Mensal_VIS.TXT_13_PG
       TXT_OBS = frm_Alt_Fic_Mensal_VIS.TXT_13_OBS
       TXT_OBS.Visible = True
    Else
       lb_dt_13.Visible = False
       lb_OBS.Visible = False
       txt_13.Visible = False
       TXT_OBS.Visible = False
        
    End If

End Sub







Private Sub TXT_CONTA_COD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{tab}"
    Else
        KeyEnter KeyCode
    End If
End Sub

Private Sub TXT_CONTA_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub TXT_DESC_CAD_KeyDown(KeyCode As Integer, Shift As Integer)
         
    If lbEmp(1).Visible = False And TXT_E_COD.Visible = False And LB_DT_EXTRA.Visible = False Then
        If KeyCode = 13 Then If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1) Then Salvar
    ElseIf LB_DT_EXTRA.Visible = True And KeyCode = 13 Then
        TXT_DT_EXTRA.SetFocus
    ElseIf KeyCode = 13 Then
        SendKeys "{tab}"
    End If
End Sub


Private Sub TXT_DESC_EXTRA_KeyDown(KeyCode As Integer, Shift As Integer)
       If KeyCode = 13 And Shift = 0 Then If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1) Then Salvar
End Sub

Private Sub TXT_DESC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{tab}"
        Pause 0.5
        If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1) Then Salvar
    Else
        KeyEnter KeyCode
    End If
End Sub

Private Sub txt_DESC_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub txt_DESC_Cad_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub



Private Sub TXT_DT_CAD_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode

End Sub

Private Sub txt_DT_CAD_Change()
GRID_EMP_DblClick
End Sub





Private Sub TXT_DT_EXTRA_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub txt_DT_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub








Private Sub txt_Emp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

       If KeyCode = 13 And Index = 1 Then If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1) Then Salvar
         Pause 0.3
       If KeyCode = 13 Then SendKeys "{TAB}"


End Sub




Private Sub txt_Logo_Cad_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode

End Sub

Private Sub TXT_LOGO_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode

End Sub

Private Sub txt_NCred_Cad_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode

End Sub

Private Sub TXT_NUM_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode

End Sub

Private Sub TXT_OP_CAD_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode

End Sub

Private Sub TXT_OP_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode

End Sub

Private Sub TXT_OP_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_OP_CAD_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_dt_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_DT_CAD_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_FUNC_CAD_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub TXT_VALOR_CAD_GotFocus()
    SendKeys "{home}+{end}"
End Sub



Private Sub TXT_VALOR_CAD_Validate(Cancel As Boolean)
    If TXT_CONTA_CAD.BoundText = "9" Then
          
           TXT_DESC_CAD = w_txt_desc & vbCrLf & "Valor : " & Format(CDbl(TXT_VALOR_CAD) - CDbl(TXT_E_JUROS), "R$ 0.00") & "    +    Juros : " & Format(TXT_E_JUROS, "R$ 0.00")

    End If
End Sub

Private Sub TXT_valor_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_VALOR_CAD_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_Conta_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_CONTA_CAD_KeyUp(KeyCode As Integer, Shift As Integer)
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
    Case 65: ' "A"
           If BarraF.Buttons("adicionar").Enabled = True Then Adicionar
    Case 69: ' "E"
           If BarraF.Buttons("editar").Enabled = True Then
                Editar
                txt_DT.SetFocus
           End If
    Case 83: ' "S"
           If BarraF.Buttons("salvar").Enabled = True Then Salvar
    Case 67: ' "C"
           If BarraF.Buttons("cancelar").Enabled = True Then Cancelar
    Case 88: ' "X"
           If BarraF.Buttons("adicionar").Enabled = True Then Excluir
    End Select
End If
End Sub





Sub CREDIARIO()
'DAO ***  p/ manipulção dos crediarios
Dim DB As dao.Database
Dim wtab As dao.Recordset
Dim wPARC As dao.Recordset


On Error GoTo loaderror
   W_LJ = Mid(frm_Alt_Fic_Mensal_VIS.TXT_CRED, 1, 2)
   w_cod = Int(Mid(frm_Alt_Fic_Mensal_VIS.TXT_CRED, 3))

    If Not IsNumeric(w_cod) Then
        MsgBox "Código de Crediário não cadastrado!", vbCritical
        GoTo sair
    End If

   If de.cnc.Execute("Select c_contrato from tab_Cred Where c_cliente = '" & W_LJ & "." & w_cod & "'").RecordCount = 0 Then

   
            'CRIA A CONEXÃO
            Set DB = DBEngine.OpenDatabase(strDirBaseCentral, False, True, "Paradox 5.x")
            
       '     Set wtab = DB.OpenRecordset("Lojb081").Clone
'            Set wtab = DB.OpenRecordset("SELECT CRED_LOJA , N_CRED , CLI_LOJA , CODIGO , VALOR_COMPRA , SALDO FROM lojb081 WHERE CLI_LOJA = '" & W_LJ & "' AND CODIGO = " & W_COD & " and SALDO > 0", dbOpenDynaset).Clone
            Set wtab = DB.OpenRecordset("SELECT CRED_LOJA , N_CRED , CLI_LOJA , CODIGO , VALOR_COMPRA , SALDO FROM lojb081 WHERE CLI_LOJA = '" & W_LJ & "' AND CODIGO = " & w_cod & " AND EXCLUIDO IS NULL", dbOpenDynaset).Clone
              
              If Not wtab.EOF Then wtab.MoveLast
              w_qtdeL = wtab.RecordCount
              wtab.MoveFirst
              'INSERI OS REGISTRO P/ O GRID
              For I = 1 To w_qtdeL
                  W_QTDE = W_QTDE + 1
                  
                  If wtab.Fields("Saldo") > 0 Then
                    'SQL - PARCELAS REFERENTE AO CONTRATO  (Q/ O SALDO SEJA MAIOR Q/ZERO)
                    Set wPARC = DB.OpenRecordset("SELECT DATA_VNC , SALDO FROM LOJB082 WHERE CRED_LOJA = '" & wtab.Fields("CRED_LOJA") & "' AND N_CRED = " & wtab.Fields("N_CRED") & " AND SALDO > 0 AND EXCLUIDO IS NULL", dbOpenDynaset)
                    
                    W_DT = ""
                    If Not wPARC.EOF Then W_DT = Format(wPARC.Fields("DATA_VNC"), "DD/MM/YYYY")
                    
                    de.cmdIncluirAuxCred (wtab.Fields("CRED_LOJA") & "." & wtab.Fields("N_CRED")), Format(W_DT, "dd/mm/yyyy"), CDbl(wtab.Fields("VALOR_COMPRA")), CDbl(wtab.Fields("SALDO")), W_LJ & "." & w_cod
                  End If
               wtab.MoveNext
              Next I
              
            DB.Close
    
    End If
    
    
    If de.rsTAB_CRED.State = 1 Then de.rsTAB_CRED.Close
    Set ADO_CRED.Recordset = de.cnc.Execute("SELECT c_contrato AS CRED, c_DT AS VCTO, c_valor AS VALOR, c_saldo AS SALDO FROM Tab_Cred WHERE C_CliENTE = '" & W_LJ & "." & w_cod & "' Order By c_DT")
    
    c_Filtro_Click
    
sair:
Exit Sub

loaderror:
    If Err.Number = 13 Then
        MsgBox "Código de Crediário não cadastrado!", vbCritical
    Else
      MsgBox Err.Number & " : " & Err.Description, vbCritical
    End If
    Resume sair
End Sub

Private Sub txtE_JUROS_Change()

End Sub
