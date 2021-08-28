VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "msCOMCTL.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.6#0"; "ACTIVETEXT.OCX"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_Alt_Desc_Calc_fixo 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALTERAÇÃO E CADASTRO DE LANÇAMENTOS PROGRAMADOS"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12675
   Icon            =   "frm_Alt_Desc_Calc_fixo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   12675
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdProcessando 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   8760
      Picture         =   "frm_Alt_Desc_Calc_fixo.frx":1042
      Style           =   1  'Graphical
      TabIndex        =   71
      TabStop         =   0   'False
      ToolTipText     =   "Processando"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.PictureBox picChecked 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   480
      Picture         =   "frm_Alt_Desc_Calc_fixo.frx":134C
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   70
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picUnchecked 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   120
      Picture         =   "frm_Alt_Desc_Calc_fixo.frx":1436
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   69
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4800
      Top             =   2040
   End
   Begin TabDlg.SSTab GUIA 
      Height          =   6375
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      Enabled         =   0   'False
      BackColor       =   12648447
      TabCaption(0)   =   "Alteração"
      TabPicture(0)   =   "frm_Alt_Desc_Calc_fixo.frx":1520
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
      Tab(0).Control(8)=   "lbEmp(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LB_EMP_DE(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lb_dt_13"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lb_OBS"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txt_13"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TXT_CONTA_cod"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txt_DT"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TXT_CONTA"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TXT_DESC"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "TXT_OP"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txt_conta_Op"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TXT_LOGO"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TXT_NUM"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txt_valor"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "TXT_E_COD_E"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txt_Obs"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txt_nficha"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "chkVisto"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "Cadastrar"
      TabPicture(1)   =   "frm_Alt_Desc_Calc_fixo.frx":153C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape2"
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(2)=   "Label7"
      Tab(1).Control(3)=   "Label10"
      Tab(1).Control(4)=   "Label11"
      Tab(1).Control(5)=   "Label12"
      Tab(1).Control(6)=   "Label8"
      Tab(1).Control(7)=   "lblogo_cad"
      Tab(1).Control(8)=   "lbncred_cad"
      Tab(1).Control(9)=   "lbEmp(2)"
      Tab(1).Control(10)=   "lbEmp(1)"
      Tab(1).Control(11)=   "lbEmp(0)"
      Tab(1).Control(12)=   "LB_EMP_D(0)"
      Tab(1).Control(13)=   "LB_EMP_D(1)"
      Tab(1).Control(14)=   "LB_EMP_D(2)"
      Tab(1).Control(15)=   "lbEmp(3)"
      Tab(1).Control(16)=   "LB_EMP_D(3)"
      Tab(1).Control(17)=   "LB_DT_EXTRA"
      Tab(1).Control(18)=   "LB_DESC_EXTRA"
      Tab(1).Control(19)=   "TXT_E_JUROS"
      Tab(1).Control(20)=   "TXT_VALOR_CAD"
      Tab(1).Control(21)=   "TXT_CONTA_Cod_CAD"
      Tab(1).Control(22)=   "TXT_OP_CAD"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "TXT_DESC_CAD"
      Tab(1).Control(24)=   "TXT_CONTA_CAD"
      Tab(1).Control(25)=   "TXT_DT_CAD"
      Tab(1).Control(26)=   "TXT_NFICHA_CAD"
      Tab(1).Control(27)=   "TXT_CONTA_CAD_op"
      Tab(1).Control(28)=   "txt_Logo_Cad"
      Tab(1).Control(29)=   "txt_NCred_Cad"
      Tab(1).Control(30)=   "txt_Emp(1)"
      Tab(1).Control(31)=   "TXT_E_COD"
      Tab(1).Control(32)=   "TXT_E_VALOR"
      Tab(1).Control(33)=   "TXT_E_SALDO"
      Tab(1).Control(34)=   "txt_Emp(2)"
      Tab(1).Control(35)=   "txt_Emp(0)"
      Tab(1).Control(36)=   "TXT_DT_EXTRA"
      Tab(1).Control(37)=   "TXT_DESC_EXTRA"
      Tab(1).ControlCount=   38
      Begin VB.CheckBox chkVisto 
         Caption         =   "VISTO"
         DataField       =   "CF_VISTO"
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
         Height          =   255
         Left            =   3960
         TabIndex        =   65
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox txt_nficha 
         Alignment       =   1  'Right Justify
         DataField       =   "CF_EMP_COD"
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
         TabIndex        =   63
         Top             =   600
         Width           =   1305
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
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   60
         Top             =   2520
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
         TabIndex        =   58
         Top             =   3210
         Visible         =   0   'False
         Width           =   3525
      End
      Begin rdActiveText.ActiveText TXT_DT_EXTRA 
         Height          =   345
         Left            =   -74760
         TabIndex        =   56
         Top             =   3330
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
         FontSize        =   8.25
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   3360
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox TXT_E_COD_E 
         Alignment       =   2  'Center
         DataField       =   "CF_EMP_COD"
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
         Left            =   240
         TabIndex        =   51
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
         Top             =   3360
         Visible         =   0   'False
         Width           =   900
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
         TabIndex        =   15
         Top             =   3360
         Visible         =   0   'False
         Width           =   540
      End
      Begin rdActiveText.ActiveText txt_valor 
         CausesValidation=   0   'False
         DataField       =   "CF_VALOR"
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
         TabIndex        =   1
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
         FontSize        =   8.25
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
         TabIndex        =   39
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
         TabIndex        =   37
         Top             =   3360
         Visible         =   0   'False
         Width           =   585
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
         TabIndex        =   35
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
         TabIndex        =   33
         Top             =   3360
         Visible         =   0   'False
         Width           =   585
      End
      Begin MSDataListLib.DataCombo TXT_CONTA_CAD_op 
         Bindings        =   "frm_Alt_Desc_Calc_fixo.frx":1558
         DataSource      =   "ADOREG"
         Height          =   360
         Left            =   -71160
         TabIndex        =   32
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
         Bindings        =   "frm_Alt_Desc_Calc_fixo.frx":1569
         DataField       =   "CF_TP_CONTA"
         DataSource      =   "ADOREG"
         Height          =   360
         Left            =   3720
         TabIndex        =   31
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
         TabIndex        =   6
         Top             =   675
         Width           =   1305
      End
      Begin VB.ComboBox TXT_OP 
         DataField       =   "CF_TP_OP"
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
         ItemData        =   "frm_Alt_Desc_Calc_fixo.frx":157A
         Left            =   4275
         List            =   "frm_Alt_Desc_Calc_fixo.frx":1587
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "+"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox TXT_DESC 
         DataField       =   "CF_DESC"
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
         TabIndex        =   4
         Top             =   2040
         Width           =   3855
      End
      Begin MSDataListLib.DataCombo TXT_CONTA 
         Bindings        =   "frm_Alt_Desc_Calc_fixo.frx":1594
         DataField       =   "CF_TP_CONTA"
         DataSource      =   "ADOREG"
         Height          =   360
         Left            =   915
         TabIndex        =   3
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
         DataField       =   "CF_DT"
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
         TabIndex        =   0
         Top             =   615
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   609
         _Version        =   393216
         Format          =   246677505
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
         TabIndex        =   7
         Top             =   690
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   609
         _Version        =   393216
         Format          =   246677505
         CurrentDate     =   38432
      End
      Begin MSDataListLib.DataCombo TXT_CONTA_CAD 
         Bindings        =   "frm_Alt_Desc_Calc_fixo.frx":15A5
         Height          =   360
         Left            =   -74085
         TabIndex        =   10
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
         TabIndex        =   11
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
         ItemData        =   "frm_Alt_Desc_Calc_fixo.frx":15B6
         Left            =   -70725
         List            =   "frm_Alt_Desc_Calc_fixo.frx":15C3
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "+"
         Top             =   2070
         Width           =   615
      End
      Begin MSDataListLib.DataCombo TXT_CONTA_cod 
         Bindings        =   "frm_Alt_Desc_Calc_fixo.frx":15D0
         DataField       =   "CF_TP_CONTA"
         DataSource      =   "ADOREG"
         Height          =   360
         Left            =   195
         TabIndex        =   2
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
         Bindings        =   "frm_Alt_Desc_Calc_fixo.frx":15E1
         Height          =   360
         Left            =   -74805
         TabIndex        =   9
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
         TabIndex        =   8
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
         FontSize        =   8.25
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
         TabIndex        =   53
         Top             =   3360
         Visible         =   0   'False
         Width           =   1020
      End
      Begin rdActiveText.ActiveText txt_13 
         Height          =   345
         Left            =   240
         TabIndex        =   61
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
         FontSize        =   8.25
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
         TabIndex        =   64
         Top             =   3000
         Visible         =   0   'False
         Width           =   3255
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
         TabIndex        =   62
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
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
         TabIndex        =   59
         Top             =   3000
         Visible         =   0   'False
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
         TabIndex        =   57
         Top             =   3120
         Visible         =   0   'False
         Width           =   735
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
         TabIndex        =   54
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
         TabIndex        =   55
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
         TabIndex        =   52
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   44
         Top             =   3000
         Visible         =   0   'False
         Width           =   570
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   38
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
         TabIndex        =   36
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
         TabIndex        =   34
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   1482
      ButtonWidth     =   1799
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
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exc. Se&l."
            Key             =   "excluirSelecionados"
            Description     =   "Excluir selecionados"
            Object.ToolTipText     =   "Excluir selecionados"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8520
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Alt_Desc_Calc_fixo.frx":15F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Alt_Desc_Calc_fixo.frx":190C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Alt_Desc_Calc_fixo.frx":1C26
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Alt_Desc_Calc_fixo.frx":1F40
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Alt_Desc_Calc_fixo.frx":225A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Alt_Desc_Calc_fixo.frx":2574
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Alt_Desc_Calc_fixo.frx":288E
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Alt_Desc_Calc_fixo.frx":49C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Alt_Desc_Calc_fixo.frx":52A2
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Alt_Desc_Calc_fixo.frx":6FAC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin Skin_Button.ctr_Button cmdSuprimir 
         Height          =   525
         Left            =   11880
         TabIndex        =   68
         TabStop         =   0   'False
         ToolTipText     =   "Suprimir confirmações de exclusão"
         Top             =   120
         Visible         =   0   'False
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "frm_Alt_Desc_Calc_fixo.frx":29376
         PICN            =   "frm_Alt_Desc_Calc_fixo.frx":29392
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin MSAdodcLib.Adodc ADOREG 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   8220
      Width           =   12675
      _ExtentX        =   22357
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
      Top             =   7890
      Visible         =   0   'False
      Width           =   12675
      _ExtentX        =   22357
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
      Bindings        =   "frm_Alt_Desc_Calc_fixo.frx":4B75C
      Height          =   6615
      Left            =   5280
      TabIndex        =   67
      Top             =   1560
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   11668
      _Version        =   393216
      Rows            =   7
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   49152
      ForeColorFixed  =   16777215
      AllowBigSelection=   0   'False
      ScrollBars      =   2
      SelectionMode   =   1
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
   Begin VB.Label lbFunc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOME DO FUNCIONÁRIO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   66
      Top             =   960
      Width           =   12855
   End
   Begin VB.Label LB_FUNC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LANÇAMENTOS PROGRAMADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   1440
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
Attribute VB_Name = "frm_Alt_Desc_Calc_fixo"
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
Dim w_unload As Boolean
Dim w_currentRow As Integer
Dim w_suprimir_exclusao

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Sub Form_Initialize()
        w_suprimir_exclusao = False
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


Public Sub AtualizarFicha()
Dim wRegFixos As Integer
    
    On Error GoTo err1
    'Atualizando lançamento automático do FIXO na Ficha do mês atual
                
                
    Dim adoFixos As ADODB.Recordset
    Dim fichaAtual As String

    fichaAtual = de.cnc.Execute("SELECT Max(M_NFICHA) FROM TAB_FICHA_MENS GROUP BY TAB_FICHA_MENS.M_F_COD HAVING (((TAB_FICHA_MENS.M_F_COD)= " & TXT_NFICHA_CAD & "))").Fields(0)

    Pause 3
    Set adoFixos = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_CODIGO = " & ADO_GRID.Recordset.Fields("CF_CODIGO")).Clone
    
    
    adoFixos.MoveFirst
    'Do While Not adoFixos.EOF
        
        de.cnc.Execute ("DELETE FROM TAB_DESC_CALC Where C_N_FICHA = " & fichaAtual & " AND C_NCRED = " & ADO_GRID.Recordset.Fields("CF_CODIGO")), wRegFixos
        de.cmdIncluirDescCalc2 txt_DT, fichaAtual, TXT_CONTA_cod, TXT_OP, txt_valor, TXT_DESC, "0", adoFixos.Fields("CF_CODIGO"), "0", "0", adoFixos.Fields("CF_EMP_COD"), 0
        MsgBox "Lançamento programado atualizado também na ficha do mês atual com sucesso!", vbOKOnly, "Atualização OK"
        
        'adoFixos.MoveNext
    'Loop
    
    'fichaAtual = Empty
    'Set adoFixos = Nothing

    
    'If wRegFixos > 0 Then MsgBox "Atualizado na Ficha Atual com sucesso!", vbExclamation
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Sub resetNavigate()
          w_adoFiltro = adoReg.Recordset.Filter
          Form_Load
          ADO_GRID.Refresh
          adoReg.Refresh
        
          adoReg.Recordset.Filter = w_adoFiltro
          ADO_GRID.Recordset.Filter = w_adoFiltro
        
          flexGRID.Refresh
        
          de.rsTAB_DESC_CALC.Close
          de.TAB_DESC_CALC
          
          formatarFlexGrid
End Sub

Private Sub flexGRID_DblClick()
    With flexGRID
        If .ColSel = 12 Then
            w_col = .Col
            .Col = 3
            If .CellPicture = picChecked Then
                Set .CellPicture = picUnchecked
            Else
                Set .CellPicture = picChecked
            End If
        End If
    End With
End Sub

Private Sub flexGRID_RowColChange()
    If flexGRID.RowSel <> 0 And flexGRID.RowSel <> w_currentRow Then
        ADO_GRID.Recordset.MoveFirst
        ADO_GRID.Recordset.Find "CF_CODIGO = " & flexGRID.TextMatrix(flexGRID.RowSel, 4), , adSearchForward
        w_currentRow = flexGRID.RowSel
    End If
End Sub

Private Sub Form_Activate()
If w_At = True Then adoReg.Refresh
    
End Sub


Sub formatarFlexGrid()
    If Not (adoReg.Recordset.EOF) Then
        flexGRID.Visible = False

        For I = 1 To flexGRID.Rows - 1
        '    If flexGRID.TextMatrix(I, 10) > 0 Then
        '         ' Varre todas as colunas da linha e seta a cor de fundo
                 For coluna = 0 To flexGRID.Cols - 1
                     flexGRID.Col = coluna
                     flexGRID.Row = I
                     flexGRID.CellBackColor = vbYellow
                     flexGRID.CellFontBold = True
                     'flexGRID_L.CellForeColor = &H80000004
                 Next coluna
        '     End If
            With flexGRID
                    .Row = I: .Col = 3: .CellPictureAlignment = 4 ' Align the checkbox
                    Set .CellPicture = picUnchecked.Picture ' Set the default checkbox picture to the empty box
            End With
        Next I
        
        flexGRID.ColAlignment(2) = flexAlignRightBottom 'valor
        flexGRID.ColAlignment(3) = flexAlignCenterBottom 'excluir
        
        flexGRID.TextMatrix(0, 0) = "Cód"
        flexGRID.TextMatrix(0, 1) = "Descrição Conta"
        flexGRID.TextMatrix(0, 2) = "Valor"
        flexGRID.TextMatrix(0, 3) = "Excluir?"
        
        flexGRID.Visible = True
        
    End If

End Sub


Private Sub Form_Load()
On Error GoTo err1

flexGRID.ColWidth(0) = 1000 'TP conta(CONTA)
flexGRID.ColWidth(1) = 4000 'conta(DESCRIÇÃO)
flexGRID.ColWidth(2) = 1300 'CF_VALOR
flexGRID.ColWidth(3) = 700 'excluir?

flexGRID.ColWidth(4) = 0
flexGRID.ColWidth(5) = 0
flexGRID.ColWidth(6) = 0
flexGRID.ColWidth(7) = 0
flexGRID.ColWidth(8) = 0
flexGRID.ColWidth(9) = 0
flexGRID.ColWidth(10) = 0
flexGRID.ColWidth(11) = 0
flexGRID.ColWidth(12) = 0

w_unload = False

'LIMPA AS CX. DE CADASTRO

    If de.rsTAB_DESC_CALC_FIXO.State = 1 Then de.rsTAB_DESC_CALC_FIXO.Close
    de.TAB_DESC_CALC_FIXO
    
    
    w_At = True
    
    
    'If frm_Alt_Funcionario.txtFCod = "" Then
        'TXT_NFICHA_CAD = TXT_NFICHA_CAD
    'Else
        TXT_NFICHA_CAD = w_CodFunc
    'End If
    TXT_CONTA_CAD = ""
    TXT_DESC_CAD = ""
    TXT_OP = ""
    TXT_DT_CAD = Date
    TXT_VALOR_CAD = 0
    
    GUIA.TabVisible(0) = True
    GUIA.TabVisible(1) = False
    
    de.rsTAB_DESC_CALC_FIXO.Requery
    
    
'sql dos Crediarios em vencto
    'w_mes = frm_Alt_Fic_Mensal_VIS.TXT_MES + 1
    'w_ano = frm_Alt_Fic_Mensal_VIS.TXT_ANO
    'W_DT = Format("01/" & w_mes & "/" & w_ano, "dd/mm/yyyy")
    'W_DT = CVDate(W_DT) + 9


    'sql registros
    
    If de.rscmdBase.State = 1 Then de.rscmdBase.Close

 If acessoTotal() Then
        de.rscmdBase.Open "SELECT * FROM TAB_DESC_CALC_FIXO  WHERE (((TAB_DESC_CALC_FIXO.CF_EMP_COD)=" & TXT_NFICHA_CAD & ")) ORDER BY TAB_DESC_CALC_FIXO.CF_VALOR, TAB_DESC_CALC_FIXO.CF_DT", , adOpenStatic, adLockOptimistic
    Else
        de.rscmdBase.Open "SELECT * FROM TAB_DESC_CALC_FIXO  WHERE (((TAB_DESC_CALC_FIXO.CF_EMP_COD)=" & TXT_NFICHA_CAD & ") AND ((TAB_DESC_CALC_FIXO.CF_TP_CONTA)<>20)) ORDER BY TAB_DESC_CALC_FIXO.CF_VALOR, TAB_DESC_CALC_FIXO.CF_DT", , adOpenStatic, adLockOptimistic
    End If

    Set adoReg.Recordset = de.rscmdBase.Clone
    de.rscmdBase.Close
    
    v_filtro = "VCTO <= #" & Format(CVDate(w_Dt), "mm/dd/YYYY") & "#"
    v_filtro_puro = ""

        
    If Not adoReg.Recordset.EOF Then
 If acessoTotal() Then
        Set ADO_GRID.Recordset = de.cnc.Execute("SELECT TAB_DESC_CALC_FIXO.CF_TP_CONTA as TP_CONTA, TAB_TP_CONTA.TP_DESC as CONTA, Format(TAB_DESC_CALC_FIXO.CF_VALOR,'R$ ###,##0.00') AS VALOR, '' as EXCLUIR, TAB_DESC_CALC_FIXO.* , TAB_TP_CONTA.TP_DESC FROM TAB_TP_CONTA, TAB_DESC_CALC_FIXO WHERE (TAB_DESC_CALC_FIXO.CF_TP_CONTA = TAB_TP_CONTA.TP_COD AND TAB_DESC_CALC_FIXO.CF_EMP_COD = " & TXT_NFICHA_CAD & " ) Order By TAB_DESC_CALC_FIXO.CF_Valor, TAB_DESC_CALC_FIXO.CF_DT").Clone
    Else
        Set ADO_GRID.Recordset = de.cnc.Execute("SELECT TAB_DESC_CALC_FIXO.CF_TP_CONTA as TP_CONTA, TAB_TP_CONTA.TP_DESC as CONTA, Format(TAB_DESC_CALC_FIXO.CF_VALOR,'R$ ###,##0.00') AS VALOR, '' as EXCLUIR, TAB_DESC_CALC_FIXO.* , TAB_TP_CONTA.TP_DESC FROM TAB_TP_CONTA, TAB_DESC_CALC_FIXO WHERE (TAB_DESC_CALC_FIXO.CF_TP_CONTA = TAB_TP_CONTA.TP_COD AND TAB_DESC_CALC_FIXO.CF_EMP_COD = " & TXT_NFICHA_CAD & " AND TAB_DESC_CALC_FIXO.CF_TP_CONTA <> 20 AND TAB_DESC_CALC_FIXO.CF_TP_CONTA <> 78) Order By TAB_DESC_CALC_FIXO.CF_Valor, TAB_DESC_CALC_FIXO.CF_DT").Clone
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
Private Sub adoReg_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1
    If Not adoReg.Recordset.EOF Then adoReg.Caption = "REGISTRO : " & adoReg.Recordset.AbsolutePosition & " / " & adoReg.Recordset.RecordCount & IIf(W_LD_FILTRO = True, " (FILTRADO)", "")
    
   If V_MOVE = True Then
        On Error Resume Next
        
       For I = 3 To 7
          
          If I > 3 And I <= 7 Then lbEmp(I).Visible = adoReg.Recordset.Fields("CF_TP_CONTA") = 31
          If I >= 3 Then txt_Emp(I).Visible = adoReg.Recordset.Fields("CF_TP_CONTA") = 31
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
        Case "editar": Editar
                       If BarraF.Buttons("editar").Enabled = False Then Me.txt_DT.SetFocus

        Case "salvar": Salvar
        Case "cancelar": Cancelar
        Case "excluir":   If BarraF.Buttons("adicionar").Enabled = True Then Excluir
        Case "excluirSelecionados": ExcluirSelecionados
        'Case "atualizar": AtualizarFicha
    End Select
End Sub

Sub ExcluirSelecionados()
Dim I, pos As Integer
Dim codigos() As String
Dim contas() As String
Dim descricoes() As String
Dim registros As String

    pos = 0
    registros = vbNewLine
    ReDim codigos(1 To flexGRID.Rows)
    ReDim contas(1 To flexGRID.Rows)
    ReDim descricoes(1 To flexGRID.Rows)
        
    With flexGRID
    For I = .Rows - 1 To 0 Step -1
        If .Rows >= 1 Then
            .Col = 3
            .Row = I
            If .CellPicture = picChecked Then
                pos = pos + 1
                codigos(pos) = flexGRID.TextMatrix(I, 4)
                contas(pos) = flexGRID.TextMatrix(I, 0)
                descricoes(pos) = flexGRID.TextMatrix(I, 1)
                registros = registros & vbNewLine & contas(pos) & ": " & descricoes(pos)
            End If
        End If
    Next I
    End With
    
    If pos > 0 Then
                
        frm_Habilitar.Show 1
        w_PSS = frm_Habilitar.txt_Pss
   
        If w_PSS = w_PassWordLib Then
                cmdProcessando.Visible = True
                If vbYes = MsgBox("Deseja realmente excluir estes " & pos & " lançamentos programados? " & registros, vbQuestion + vbYesNo) Then
                    For I = pos To 1 Step -1
                        excluirLanc codigos(I)
                    Next I
                    MsgBox "Os " & pos & " registros foram excluídos com sucesso!", vbOKOnly
                    resetNavigate
                Else
                    MsgBox "Operação cancelada.", vbOKOnly
                End If
                cmdProcessando.Visible = False
        Else
            MsgBox "Senha de Liberação Incorreta!", vbCritical
        End If
    Else
        MsgBox "Selecione pelo menos 1 registro para a exclusão.", vbInformation, "Operação cancelada"
    End If
    
End Sub


Private Sub excluirLanc(ByVal codigo As String)
On Error GoTo err1
Dim ficha, emp_cod As Integer
          
        
Set adoExcluir = de.cnc.Execute("SELECT TAB_DESC_CALC_FIXO.CF_EMP_COD AS emp_cod, " _
    & "(SELECT Max(M_NFICHA) FROM TAB_FICHA_MENS GROUP BY TAB_FICHA_MENS.M_F_COD HAVING (((TAB_FICHA_MENS.M_F_COD)= TAB_DESC_CALC_FIXO.CF_EMP_COD))) as ficha " _
    & "FROM TAB_DESC_CALC_FIXO WHERE (((TAB_DESC_CALC_FIXO.CF_CODIGO)=" & codigo & "))").Clone

    ficha = adoExcluir.Fields("ficha")
    emp_cod = adoExcluir.Fields("emp_cod")
        
    de.cnc.Execute ("DELETE FROM TAB_DESC_CALC_FIXO Where CF_CODIGO = " & codigo)
    de.cnc.Execute ("DELETE FROM TAB_DESC_CALC Where C_N_FICHA = " & ficha & " AND C_NCRED = " & codigo)
   
sair:
    Exit Sub
err1:
    If Not Err.Number = -2147467259 Then
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    Else
        MsgBox Err.Number & " : " & Err.Description, vbCritical
        'MsgBox "NÃO É POSSÍVEL EXCLUIR ESTA FICHA MENSAL, DEVIDO A CÁLCULOS RELACIONADAS A ELA!", vbCritical
        adoReg.Refresh
    End If
    Resume sair

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
    
    TXT_NFICHA_CAD = TXT_NFICHA_CAD
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

On Error GoTo errPula

If Not ADO_GRID.Recordset.State = 0 Then

    If ADO_GRID.Recordset.Fields("CF_VISTO") = True And w_PSS = "" And w_unload = False Then
    frm_Habilitar.Show 1
    w_PSS = frm_Habilitar.txt_Pss
    Else
errPula:
    
    w_PSS = w_PassWordLib
    End If
End If

    


On Error GoTo err1

If w_PSS = w_PassWordLib Then
    

    w_At = False
    Pause 0.5
'SE FOR CREDIARIO MOSTRA O GRID DE CREDIARIOS
     If TXT_CONTA.BoundText = "17" And GUIA.TabVisible(0) = True Then
        MsgBox "Não é permitido editar Crediário, exclua e adicione novamente!", vbCritical
        GoTo sair
     'ElseIf (TXT_CONTA.BoundText = "31" Or TXT_CONTA.BoundText = "9") And GUIA.TabVisible(0) = True Then
     '   MsgBox "Não é permitido editar Empréstimo, exclua e adicione novamente!", vbCritical
     '   GoTo sair
'     ElseIf TXT_CONTA.BoundText = "32" And GUIA.TabVisible(0) = True Then
 '       MsgBox "Não é permitido editar 13º, exclua e adicione novamente!", vbCritical
'        GoTo sair
     ElseIf TXT_CONTA.BoundText = "24" And GUIA.TabVisible(0) = True Then
        MsgBox "Não é permitido editar Férias, exclua e adicione novamente!", vbCritical
        GoTo sair
        
     Else
        'GRID_CRED.Visible = False
        'c_Filtro.Visible = False
     End If
    
    BarraF.Buttons("salvar").Enabled = Not BarraF.Buttons("salvar").Enabled
    BarraF.Buttons("cancelar").Enabled = Not BarraF.Buttons("cancelar").Enabled
    BarraF.Buttons("editar").Enabled = Not BarraF.Buttons("editar").Enabled
    BarraF.Buttons("adicionar").Enabled = Not BarraF.Buttons("adicionar").Enabled
    'BarraF.Buttons("atualizar").Enabled = Not BarraF.Buttons("atualizar").Enabled
    
    'GRID_CRED.Enabled = Not GRID_CRED.Enabled
    flexGRID.Enabled = Not flexGRID.Enabled
        
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
Dim codFixo As String

Dim fichaAtual As String

        
'If ADO_GRID.Recordset.Fields("CF_VISTO") = True And w_PSS = "" Then
    frm_Habilitar.Show 1
    w_PSS = frm_Habilitar.txt_Pss
'Else
'    w_PSS = w_PassWordLib
'End If

If w_PSS = w_PassWordLib Then
        
    If vbYes = MsgBox("DESEJA REALMENTE EXCLUIR O LANÇAMENTO PROGRAMADO (" & TXT_CONTA & ")?", vbQuestion + vbYesNo) Then
        
        ' Parametros para exclusão do fixo lá na ficha mensal atual
        fichaAtual = de.cnc.Execute("SELECT Max(M_NFICHA) FROM TAB_FICHA_MENS GROUP BY TAB_FICHA_MENS.M_F_COD HAVING (((TAB_FICHA_MENS.M_F_COD)= " & TXT_NFICHA_CAD & "))").Fields(0)
        codFixo = ADO_GRID.Recordset.Fields("CF_CODIGO")
        
       
          
        w_Conta = TXT_CONTA.BoundText
     
        
        
        
        '*** Exclui o registro
        adoReg.Recordset.Find "CF_CODIGO = " & ADO_GRID.Recordset.Fields("CF_CODIGO") & ""
        W_POS = adoReg.Recordset.AbsolutePosition - 1
        adoReg.Recordset.Delete
        w_adoFiltro = adoReg.Recordset.Filter
        Form_Load
        ADO_GRID.Refresh
        adoReg.Refresh
      
        adoReg.Recordset.Filter = w_adoFiltro
        ADO_GRID.Recordset.Filter = w_adoFiltro
      
        flexGRID.Refresh
      
        de.rsTAB_DESC_CALC_FIXO.Close
        de.TAB_DESC_CALC_FIXO
        '*** CALCULA O TOTAL - APÓS O NOVO VALOR ***
        
     'If acessoTotal() Then
     '       W_MAIS = de.cnc.Execute("SELECT SUM(CF_VALOR) AS MAIS FROM TAB_DESC_CALC_FIXO  WHERE (CF_TP_OP = '+') AND (CF_EMP_COD = " & TXT_NFICHA_CAD & ")").Fields("MAIS")
     '       W_MENOS = de.cnc.Execute("SELECT SUM(CF_VALOR) AS MENOS FROM TAB_DESC_CALC_FIXO WHERE (CF_TP_OP = '-') AND (CF_EMP_COD = " & TXT_NFICHA_CAD & ")").Fields("MENOS")
     '   Else
     '       W_MAIS = de.cnc.Execute("SELECT SUM(CF_VALOR) AS MAIS FROM TAB_DESC_CALC_FIXO  WHERE (CF_TP_OP = '+') AND (CF_EMP_COD = " & TXT_NFICHA_CAD & ") AND (TAB_DESC_CALC_FIXO.CF_TP_CONTA <> 20)").Fields("MAIS")
     '       W_MENOS = de.cnc.Execute("SELECT SUM(CF_VALOR) AS MENOS FROM TAB_DESC_CALC_FIXO WHERE (CF_TP_OP = '-') AND (CF_EMP_COD = " & TXT_NFICHA_CAD & ") AND (TAB_DESC_CALC_FIXO.CF_TP_CONTA <> 20)").Fields("MENOS")
     '   End If
     '
     '   W_TOTAL = IIf(IsNull(W_MAIS), 0, W_MAIS) + IIf(IsNull(W_MENOS), 0, W_MENOS)
     '
     '   de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_TOTAL = '" & CDbl(W_TOTAL) & "' WHERE (M_NFICHA = " & TXT_NFICHA_CAD & ")"
      
         w_PSS = ""
      '  Grid.ReBind
      
      
    '*** ATUALIZA A ULTIMA DATA DE PAGAMENTO NA TAB_EMPRESTIMO *** SE FOR DESCONTO
    'If w_Conta = "14" Then
   '
   '     '*** Atualiza VALOR DO SALDO DEVEDOR EM TAB_FUNCIONARIO ***
   '     de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_SALDO_ANT = F_SALDO_ANT + '" & CDbl(w_Valor) & "' WHERE (F_Codigo = " & frm_Alt_Fic_Mensal_VIS.adoReg.Recordset.Fields("m_F_COD") & ")"
   '
   ' End If
        
  
    
    Dim adoFixos As ADODB.Recordset
    
    'Sleep (1000)
    Set adoFixos = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_CODIGO = " & codFixo).Clone

    'Do While Not adoFixos.EOF
        de.cnc.Execute ("DELETE FROM TAB_DESC_CALC Where C_N_FICHA = " & fichaAtual & " AND C_NCRED = " & codFixo)
        'adoFixos.MoveNext
    'Loop
    
    fichaAtual = Empty
    Set adoFixos = Nothing
    
      
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
        MsgBox Err.Number & " : " & Err.Description, vbCritical
        'MsgBox "NÃO É POSSÍVEL EXCLUIR ESTA FICHA MENSAL, DEVIDO A CÁLCULOS RELACIONADAS A ELA!", vbCritical
        adoReg.Refresh
    End If
    Resume sair
End Sub

Private Sub Fechar()
On Error GoTo err1
        
       

        
        '*** CALCULA O TOTAL - APÓS O NOVO VALOR ***
        'If UCase(w_usuario) <> "KELLEN" Then
            'W_MAIS = de.cnc.Execute("SELECT SUM(CF_VALOR) AS MAIS FROM TAB_DESC_CALC_FIXO  WHERE (CF_TP_OP = '+') AND (CF_EMP_COD = " & TXT_NFICHA_CAD & ")").Fields("MAIS")
            'W_MENOS = de.cnc.Execute("SELECT SUM(CF_VALOR) AS MENOS FROM TAB_DESC_CALC_FIXO WHERE (CF_TP_OP = '-') AND (CF_EMP_COD = " & TXT_NFICHA_CAD & ")").Fields("MENOS")
        'Else
            'W_MAIS = de.cnc.Execute("SELECT SUM(CF_VALOR) AS MAIS FROM TAB_DESC_CALC_FIXO  WHERE (CF_TP_OP = '+') AND (CF_EMP_COD = " & TXT_NFICHA_CAD & ") AND (TAB_DESC_CALC_FIXO.CF_TP_CONTA <> 20)").Fields("MAIS")
            'W_MENOS = de.cnc.Execute("SELECT SUM(CF_VALOR) AS MENOS FROM TAB_DESC_CALC_FIXO WHERE (CF_TP_OP = '-') AND (CF_EMP_COD = " & TXT_NFICHA_CAD & ") AND (TAB_DESC_CALC_FIXO.CF_TP_CONTA <> 20)").Fields("MENOS")
        'End If
        'W_MAIS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MAIS FROM TAB_DESC_CALC_FIXO  WHERE (C_TP_OP = '+') AND (CF_EMP_COD = " & TXT_NFICHA_CAD & ")").Fields("MAIS")
        'W_MENOS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MENOS FROM TAB_DESC_CALC_FIXO WHERE (C_TP_OP = '-') AND (CF_EMP_COD = " & TXT_NFICHA_CAD & ")").Fields("MENOS")
        
       ' W_TOTAL = IIf(IsNull(W_MENOS), 0, W_MENOS) + IIf(IsNull(W_MAIS), 0, W_MAIS)
        
        
        '***Pega saldo de emprestimo e atualiza ***
        'w_Saldo_Emp = de.cnc.Execute("Select SUM(e_Saldo) as Saldo from Tab_Emprestimo where e_F_Cod = " & frm_Alt_Fic_Mensal_VIS.txt_F_COD & "").Fields(0)
        'w_Saldo_Emp = IIf(IsNull(w_Saldo_Emp), 0, w_Saldo_Emp)
        
      
        '*** ATUALIZA NA TAB FICHA
        'de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_TOTAL = '" & CDbl(W_TOTAL) & "' WHERE (M_NFICHA = " & TXT_NFICHA_CAD & ")"
        '*** ATUALIZA SALDO DO EMPRESTIMO NO CAD. DE FUNC.
        'de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_EMPRESTIMO = '" & CDbl(w_Saldo_Emp) & "' WHERE (F_CODIGO = " & frm_Alt_Fic_Mensal_VIS.txt_F_COD & ")"
       
        
        'de.rsTAB_DESC_CALC_FIXO.Requery
        'de.rsTAB_DESC_CALC_FIXO.Close
        'de.TAB_DESC_CALC_FIXO
        
        'If de.rsTAB_FICHA_MENS.State = 1 Then de.rsTAB_FICHA_MENS.Requery
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
Dim db As dao.Database
Dim wtab As dao.Recordset
Dim wPARC As dao.Recordset
Dim incluirFixo As Boolean

incluirFixo = False
              
              
If GUIA.TabVisible(0) = True Then   '****   ALTERAR   ****
        incluirFixo = False
    
        If (txt_valor < 0 And (TXT_OP = "+" Or TXT_OP = "=")) Or (txt_valor > 0 And TXT_OP = "-") Then txt_valor = txt_valor * -1
        'If (txt_valor < 0 And TXT_OP = "+") Or (txt_valor > 0 And TXT_OP = "-") Then txt_valor = txt_valor * -1
        'AtualizarFicha
        adoReg.Recordset.UpdateBatch adAffectCurrent
        
        'If TXT_CONTA.BoundText = 32 Then
            'frm_Alt_Fic_Mensal_VIS.TXT_13_OBS = TXT_OBS
            'frm_Alt_Fic_Mensal_VIS.TXT_13_PG = txt_13
            '*** Atualiza Dt 13º ***  TAB_FUNCIONARIO
        '    de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_13_OK = 0 , F_13_PG = '" & txt_13 & "' , F_13_OBS = '" & txt_Obs & "' WHERE (F_Codigo = " & frm_Alt_Fic_Mensal_VIS.ADOREG.Recordset.Fields("M_F_COD") & ")"
            '*** Atualiza Dt 13º***   TAB_FICHA_MENS
            'de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_13_OK = 0 , M_13_PG = '" & txt_13 & "', M_13_OBS = '" & txt_Obs & "'  WHERE (M_F_Cod = " & frm_Alt_Fic_Mensal_VIS.ADOREG.Recordset.Fields("M_F_COD") & ")"
        'End If
        
               
        Editar
        w_At = True
            
        'Pause 1
        adoReg.Refresh
        
        de.rsTAB_DESC_CALC_FIXO.Requery
        If de.rscmdBase.State = 1 Then de.rscmdBase.Close
    
        
     If acessoTotal() Then
         de.rscmdBase.Open "SELECT * FROM TAB_DESC_CALC_FIXO  WHERE (((TAB_DESC_CALC_FIXO.CF_EMP_COD)=" & TXT_NFICHA_CAD & ")) ORDER BY TAB_DESC_CALC_FIXO.CF_VALOR, TAB_DESC_CALC_FIXO.CF_DT", , adOpenStatic, adLockOptimistic
     Else
         de.rscmdBase.Open "SELECT * FROM TAB_DESC_CALC_FIXO  WHERE (((TAB_DESC_CALC_FIXO.CF_EMP_COD)=" & TXT_NFICHA_CAD & ") AND ((TAB_DESC_CALC_FIXO.CF_TP_CONTA)<>20)) ORDER BY TAB_DESC_CALC_FIXO.CF_VALOR, TAB_DESC_CALC_FIXO.CF_DT", , adOpenStatic, adLockOptimistic
     End If
    
    
     Set adoReg.Recordset = de.rscmdBase.Clone
        
        
        'AtualizarFicha
        'Pause 1
        'AtualizarFicha
        
        
        adoReg.Refresh
        
     If acessoTotal() Then
         Set ADO_GRID.Recordset = de.cnc.Execute("SELECT TAB_DESC_CALC_FIXO.CF_TP_CONTA as TP_CONTA, TAB_TP_CONTA.TP_DESC as CONTA, Format(TAB_DESC_CALC_FIXO.CF_VALOR,'R$ ###,##0.00') AS VALOR, '' as EXCLUIR, TAB_DESC_CALC_FIXO.* , TAB_TP_CONTA.TP_DESC FROM TAB_TP_CONTA, TAB_DESC_CALC_FIXO WHERE ( TAB_DESC_CALC_FIXO.CF_TP_CONTA = TAB_TP_CONTA.TP_COD AND TAB_DESC_CALC_FIXO.CF_EMP_COD = " & adoReg.Recordset.Fields("CF_EMP_COD") & " ) Order By TAB_DESC_CALC_FIXO.CF_Valor, TAB_DESC_CALC_FIXO.CF_DT").Clone
     Else
         Set ADO_GRID.Recordset = de.cnc.Execute("SELECT TAB_DESC_CALC_FIXO.CF_TP_CONTA as TP_CONTA, TAB_TP_CONTA.TP_DESC as CONTA, Format(TAB_DESC_CALC_FIXO.CF_VALOR,'R$ ###,##0.00') AS VALOR, '' as EXCLUIR, TAB_DESC_CALC_FIXO.* , TAB_TP_CONTA.TP_DESC FROM TAB_TP_CONTA, TAB_DESC_CALC_FIXO WHERE ( TAB_DESC_CALC_FIXO.CF_TP_CONTA = TAB_TP_CONTA.TP_COD AND TAB_DESC_CALC_FIXO.CF_EMP_COD = " & adoReg.Recordset.Fields("CF_EMP_COD") & " ) AND (TAB_DESC_CALC_FIXO.CF_TP_CONTA <> 20) AND (TAB_DESC_CALC_FIXO.CF_TP_CONTA <> 78) Order By TAB_DESC_CALC_FIXO.CF_Valor, TAB_DESC_CALC_FIXO.CF_DT").Clone
     End If

        ADO_GRID.Refresh
        'Pause 1
        'FlexGrid.ReBind
        
        resetNavigate
    
        W_FICHA = TXT_NFICHA
        w_PSS = ""
        
        'AtualizarFicha
    
Else    '**** CADASTRAR ****

    incluirFixo = True
    
    If lb_form = "mensal" Then
        'w_mes = frm_Alt_Fic_Mensal_VIS.TXT_MES
        'w_ano = frm_Alt_Fic_Mensal_VIS.TXT_ANO
    ElseIf lb_form = "visualizar" Then
        'w_mes = frm_Alt_Fic_Mensal_VIS.ADOREG.Recordset.Fields("M_MES")
        'w_ano = frm_Alt_Fic_Mensal_VIS.TXT_ANO
    End If
    
    If TXT_DT_EXTRA.Visible = True And TXT_DT_EXTRA = "" Then
        MsgBox "Você deve preencher Data do pagamento!  " & LB_DT_EXTRA, vbInformation
        TXT_DT_EXTRA.SetFocus
        GoTo sair
    End If
    
    If (CDbl(Format(TXT_DT_CAD, "mm")) >= CDbl(w_mes) Or CDbl(Format(TXT_DT_CAD, "mm")) = CDbl(w_mes) - 1) And (Not CDbl(TXT_VALOR_CAD) = 0 Or TXT_VALOR_CAD <> "") And TXT_OP_CAD <> "" Then
        
        
        If (TXT_VALOR_CAD < 0 And TXT_OP_CAD = "+") Or (TXT_VALOR_CAD > 0 And TXT_OP_CAD = "-") Then TXT_VALOR_CAD = TXT_VALOR_CAD * -1
        
        If txt_Logo_Cad = "" And TXT_CONTA_CAD.BoundText <> "31" Then   '*** NÃO SEJA CREDIARIOS E NÃO SEJA EMPRESTIMO ***
                de.cmdIncluirDescCalcFixo TXT_DT_CAD, TXT_NFICHA_CAD, TXT_CONTA_CAD.BoundText, TXT_OP_CAD, TXT_VALOR_CAD, IIf(TXT_DESC_CAD = "", " ", TXT_DESC_CAD)
        ElseIf TXT_CONTA_CAD.BoundText <> "31" Then  '*** DIFERENTE DE EMPRESTIMO ***
                de.cmdIncluirDescCalcFixo TXT_DT_CAD, TXT_NFICHA_CAD, TXT_CONTA_CAD.BoundText, TXT_OP_CAD, TXT_VALOR_CAD, TXT_DESC_CAD
        
        End If
        
        'para não atualizar o mês atual
        'incluirFixo = True
                
        MsgBox "Registro salvo com sucesso!", vbInformation
        
        TXT_CONTA_CAD = ""
        
        BarraF.Buttons("excluir").Enabled = True
        Editar
        Form_Load
        
        W_FICHA = TXT_NFICHA_CAD
        
        
    ElseIf Not (CDbl(Format(TXT_DT_CAD, "mm")) = CDbl(frm_Alt_Fic_Mensal_VIS.TXT_MES) Or CDbl(Format(TXT_DT_CAD, "mm")) = CDbl(frm_Alt_Fic_Mensal_VIS.TXT_MES) - 1) Then
        MsgBox "Só é permitido data do mês anterior ou do mês atual!", vbExclamation
    Else
        MsgBox "Preencha os Campos!", vbCritical
    End If


End If

   If TXT_NFICHA <> "" And incluirFixo = True Then
        
        de.rsTAB_DESC_CALC_FIXO.Close
        de.TAB_DESC_CALC_FIXO
        
        '*** CALCULA O TOTAL - APÓS O NOVO VALOR ***
        W_MAIS = de.cnc.Execute("SELECT SUM(CF_VALOR) AS MAIS FROM TAB_DESC_CALC_FIXO  WHERE (CF_TP_OP = '+') AND (CF_EMP_COD = " & TXT_NFICHA_CAD & ")").Fields("MAIS")
        W_MENOS = de.cnc.Execute("SELECT SUM(CF_VALOR) AS MENOS FROM TAB_DESC_CALC_FIXO WHERE (CF_TP_OP = '-') AND (CF_EMP_COD = " & TXT_NFICHA_CAD & ")").Fields("MENOS")
        
        W_TOTAL = IIf(IsNull(W_MAIS), 0, W_MAIS) - IIf(IsNull(W_MENOS), 0, W_MENOS)

        de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_TOTAL = '" & CDbl(W_TOTAL) & "' WHERE (M_NFICHA = " & TXT_NFICHA & ")"
            
  End If
  
  If incluirFixo = True Then
    'Incluindo lançamento automático do FIXO na Ficha do mês atual
                
    Dim adoFixos As ADODB.Recordset
    
    Dim fichaAtual As String
    Dim ultimoFixo As String

    fichaAtual = de.cnc.Execute("SELECT Max(M_NFICHA) FROM TAB_FICHA_MENS GROUP BY TAB_FICHA_MENS.M_F_COD HAVING (((TAB_FICHA_MENS.M_F_COD)= " & TXT_NFICHA_CAD & "))").Fields(0)
    ultimoFixo = de.cnc.Execute("SELECT Max([CF_CODIGO]) FROM TAB_DESC_CALC_FIXO").Fields(0)

    Set adoFixos = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_CODIGO = " & ultimoFixo).Clone

    'Do While Not adoFixos.EOF
        de.cmdIncluirDescCalc2 Date, fichaAtual, adoFixos.Fields("CF_TP_CONTA"), adoFixos.Fields("CF_TP_OP"), adoFixos.Fields("CF_VALOR"), adoFixos.Fields("CF_DESC"), "0", adoFixos.Fields("CF_CODIGO"), "0", "0", adoFixos.Fields("CF_EMP_COD"), 0
        'adoFixos.MoveNext
    'Loop
    
    fichaAtual = Empty
    ultimoFixo = Empty
    Set adoFixos = Nothing
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





















Private Sub GRID_CRED_DblClick()
    mnuSelSel_Click
End Sub

Private Sub GRID_CRED_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
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





Private Sub Form_Unload(Cancel As Integer)

Fechar

End Sub

Private Sub GUIA_GotFocus()
    If GUIA.Tab = 0 Then
        txt_DT.SetFocus
    Else
        TXT_DT_CAD.SetFocus
    End If
End Sub



Private Sub mnuSuprimir_Click(Index As Integer)
    If w_suprimir_exclusao Then
        w_suprimir_exclusao = False
    Else
        w_suprimir_exclusao = True
    End If
    cmdSuprimir.Visible = w_suprimir_exclusao
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
        TXT_OP_CAD = TXT_CONTA_CAD_op.text
       w_At = False
   End If

        '** Sql EMP. P/ GRID
        'If TXT_CONTA_Cod_CAD.BoundText = "9" Then Set ado_EMP.Recordset = de.cnc.Execute("SELECT * FROM TAB_EMPRESTIMO WHERE E_F_COD = " & frm_Alt_Fic_Mensal_VIS.TXT_FUNC.BoundText & " AND E_SALDO > 0").Clone
            
         'GRID_EMP.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
    
         'LB_EMP_D(0).Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         'LB_EMP_D(1).Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         'LB_EMP_D(2).Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         'LB_EMP_D(3).Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         'TXT_E_COD.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         'TXT_E_SALDO.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         'TXT_E_VALOR.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         'TXT_E_JUROS.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         'c_Filtro_Emp.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"

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
        
        
    'ElseIf TXT_CONTA_CAD.BoundText = "9" Then  '*** DESCONTO DO EMPRESTIMO ***
        'Emprestimo descontos
        
    End If
    
    If UCase(TXT_CONTA_CAD.text) = "DESC. CRED." Then
        
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
        'GRID_CRED.Visible = False
        'c_Filtro.Visible = False
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
    If KeyCode = 13 Then Sendkeys "{tab}"
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
            
           If adoReg.Recordset.Fields("CF_TP_CONTA") <> TXT_CONTA.BoundText Then
                txt_conta_Op.BoundText = TXT_CONTA.BoundText
                
                TXT_CONTA_cod.BoundText = TXT_CONTA.BoundText
                TXT_OP = txt_conta_Op.text
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
            'LBNCRED.Visible = False
            'LBLOGO.Visible = False
            'c_Filtro.Visible = False
            'GRID_CRED.Visible = False
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
        'If TXT_CONTA_Cod_CAD.BoundText = "9" Then Set ado_EMP.Recordset = de.cnc.Execute("SELECT * FROM TAB_EMPRESTIMO WHERE E_F_COD = " & frm_Alt_Fic_Mensal_VIS.TXT_FUNC.BoundText & " AND E_SALDO > 0").Clone
            
         'GRID_EMP.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
    
         'LB_EMP_D(0).Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         'LB_EMP_D(1).Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         'LB_EMP_D(2).Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         'LB_EMP_D(3).Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         'TXT_E_COD.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         'TXT_E_SALDO.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         'TXT_E_VALOR.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         'TXT_E_JUROS.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         ''c_Filtro_Emp.Visible = TXT_CONTA_Cod_CAD.BoundText = "9"
         
         
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
    If KeyCode = 13 Then Sendkeys "{tab}"

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
    
    'TXT_E_COD_E.Visible = (TXT_CONTA.BoundText = "9" Or TXT_CONTA.BoundText = "31")
    'LB_EMP_DE(0).Visible = (TXT_CONTA.BoundText = "9" Or TXT_CONTA.BoundText = "31")
    'TXT_E_JUROS_E.Visible = (TXT_CONTA.BoundText = "9")
    'LB_EMP_DE(1).Visible = (TXT_CONTA.BoundText = "9")
    
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
       'lb_dt_13.Visible = False
       lb_OBS.Visible = False
       'txt_13.Visible = False
       TXT_OBS.Visible = False
        
    End If

End Sub







Private Sub TXT_CONTA_COD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Sendkeys "{tab}"
    Else
        KeyEnter KeyCode
    End If
End Sub

Private Sub TXT_CONTA_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{tab}"

End Sub

Private Sub TXT_DESC_CAD_KeyDown(KeyCode As Integer, Shift As Integer)
         
    If lbEmp(1).Visible = False And TXT_E_COD.Visible = False And LB_DT_EXTRA.Visible = False Then
        If KeyCode = 13 Then If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1) Then Salvar
    ElseIf LB_DT_EXTRA.Visible = True And KeyCode = 13 Then
        TXT_DT_EXTRA.SetFocus
    ElseIf KeyCode = 13 Then
        Sendkeys "{tab}"
    End If
End Sub


Private Sub TXT_DESC_EXTRA_KeyDown(KeyCode As Integer, Shift As Integer)
       If KeyCode = 13 And Shift = 0 Then If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1) Then Salvar
End Sub

Private Sub TXT_DESC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Sendkeys "{tab}"
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
       If KeyCode = 13 Then Sendkeys "{TAB}"


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
    Sendkeys "{home}+{end}"
End Sub



Private Sub TXT_VALOR_CAD_Validate(Cancel As Boolean)
    'If TXT_CONTA_CAD.BoundText = "9" Then
          
    '       TXT_DESC_CAD = w_txt_desc & vbCrLf & "Valor : " & Format(CDbl(TXT_VALOR_CAD) - CDbl(TXT_E_JUROS), "R$ 0.00") & "    +    Juros : " & Format(TXT_E_JUROS, "R$ 0.00")

    'End If
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
Dim db As dao.Database
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
            Set db = DBEngine.OpenDatabase(strDirBaseCentral, False, True, "Paradox 5.x")
            
       '    Set wtab = DB.OpenRecordset("Lojb081").Clone
'           Set wtab = DB.OpenRecordset("SELECT CRED_LOJA , N_CRED , CLI_LOJA , CODIGO , VALOR_COMPRA , SALDO FROM lojb081 WHERE CLI_LOJA = '" & W_LJ & "' AND CODIGO = " & W_COD & " and SALDO > 0", dbOpenDynaset).Clone
            Set wtab = db.OpenRecordset("SELECT CRED_LOJA , N_CRED , CLI_LOJA , CODIGO , VALOR_COMPRA , SALDO FROM lojb081 WHERE CLI_LOJA = '" & W_LJ & "' AND CODIGO = " & w_cod & " AND EXCLUIDO IS NULL", dbOpenDynaset).Clone
              
              If Not wtab.EOF Then wtab.MoveLast
              w_qtdeL = wtab.RecordCount
              wtab.MoveFirst
              'INSERI OS REGISTRO P/ O GRID
              For I = 1 To w_qtdeL
                  W_QTDE = W_QTDE + 1
                  
                  If wtab.Fields("Saldo") > 0 Then
                    'SQL - PARCELAS REFERENTE AO CONTRATO  (Q/ O SALDO SEJA MAIOR Q/ZERO)
                    Set wPARC = db.OpenRecordset("SELECT DATA_VNC , SALDO FROM LOJB082 WHERE CRED_LOJA = '" & wtab.Fields("CRED_LOJA") & "' AND N_CRED = " & wtab.Fields("N_CRED") & " AND SALDO > 0 AND EXCLUIDO IS NULL", dbOpenDynaset)
                    
                    w_Dt = ""
                    If Not wPARC.EOF Then w_Dt = Format(wPARC.Fields("DATA_VNC"), "DD/MM/YYYY")
                    
                    de.cmdIncluirAuxCred (wtab.Fields("CRED_LOJA") & "." & wtab.Fields("N_CRED")), Format(w_Dt, "dd/mm/yyyy"), CDbl(wtab.Fields("VALOR_COMPRA")), CDbl(wtab.Fields("SALDO")), W_LJ & "." & w_cod
                  End If
               wtab.MoveNext
              Next I
              
            db.Close
    
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
