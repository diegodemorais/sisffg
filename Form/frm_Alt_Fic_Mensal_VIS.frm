VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.6#0"; "ACTIVETEXT.OCX"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_Alt_Fic_Mensal_VIS 
   BackColor       =   &H80000000&
   Caption         =   "FICHA MENSAL"
   ClientHeight    =   9750
   ClientLeft      =   5370
   ClientTop       =   3210
   ClientWidth     =   15450
   Icon            =   "frm_Alt_Fic_Mensal_VIS.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   15450
   Begin VB.CommandButton cmdDelFichas 
      BackColor       =   &H8000000D&
      Caption         =   "Deletar Fichas ñ exist."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaskColor       =   &H00FF0000&
      TabIndex        =   164
      Top             =   7920
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdDtDEM 
      BackColor       =   &H8000000D&
      Caption         =   "Atualizar Dt Demissão"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaskColor       =   &H00FF0000&
      TabIndex        =   163
      Top             =   7560
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdFixoX 
      BackColor       =   &H8000000D&
      Caption         =   "Gerar Fixo específico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8640
      MaskColor       =   &H00FF0000&
      TabIndex        =   154
      Top             =   7320
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdSindical 
      BackColor       =   &H8000000D&
      Caption         =   "Gerar Contr. Sindical"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8640
      MaskColor       =   &H00FF0000&
      TabIndex        =   153
      Top             =   6960
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdComerciario 
      BackColor       =   &H8000000D&
      Caption         =   "Gerar Comerciários"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaskColor       =   &H00FF0000&
      TabIndex        =   151
      Top             =   6600
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdBonus 
      BackColor       =   &H8000000D&
      Caption         =   "Exportar Custo Bônus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaskColor       =   &H00FF0000&
      TabIndex        =   162
      Top             =   6240
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdDespesa99 
      BackColor       =   &H8000000D&
      Caption         =   "Exportar Despesas RP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8640
      MaskColor       =   &H00FF0000&
      TabIndex        =   156
      Top             =   5970
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdDespesa 
      BackColor       =   &H8000000D&
      Caption         =   "Despesas CX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8640
      MaskColor       =   &H00FF0000&
      TabIndex        =   155
      Top             =   5640
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdExportarContas 
      BackColor       =   &H8000000D&
      Caption         =   "Exportar p/ Contas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaskColor       =   &H00FF0000&
      TabIndex        =   152
      Top             =   5280
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdFaltas 
      Height          =   435
      Left            =   4680
      Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":1042
      Style           =   1  'Graphical
      TabIndex        =   149
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CheckBox ck_pg_vt 
      Caption         =   "Check1"
      DataField       =   "M_PG_VT"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4320
      TabIndex        =   147
      TabStop         =   0   'False
      Top             =   3840
      Width           =   195
   End
   Begin VB.CommandButton cmdFixos 
      BackColor       =   &H8000000D&
      Caption         =   "Gerar Fixos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaskColor       =   &H00FF0000&
      TabIndex        =   146
      Top             =   4920
      Visible         =   0   'False
      Width           =   2000
   End
   Begin Skin_Button.ctr_Button cmdSalarioCX 
      Height          =   285
      Left            =   6120
      TabIndex        =   145
      TabStop         =   0   'False
      ToolTipText     =   "Salários CXs"
      Top             =   1320
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   503
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   12632319
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_Alt_Fic_Mensal_VIS.frx":134C
      PICN            =   "frm_Alt_Fic_Mensal_VIS.frx":1368
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton cmdEmprestimo 
      BackColor       =   &H8000000D&
      Caption         =   "Gerar Baixas Emprest."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaskColor       =   &H00FF0000&
      TabIndex        =   143
      Top             =   4560
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdAddLanç_SalFTodos 
      BackColor       =   &H8000000D&
      Caption         =   "Gerar Salários Família"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaskColor       =   &H00FF0000&
      TabIndex        =   121
      Top             =   4200
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdSaldo 
      BackColor       =   &H8000000D&
      Caption         =   "Gerar Saldos Anterior"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaskColor       =   &H00FF0000&
      TabIndex        =   108
      Top             =   3840
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmd13 
      BackColor       =   &H8000000D&
      Caption         =   "Gerar 13ºs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaskColor       =   &H00FF0000&
      TabIndex        =   115
      Top             =   3480
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdComissGerente 
      BackColor       =   &H8000000D&
      Caption         =   "Comiss. Gerente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaskColor       =   &H00FF0000&
      TabIndex        =   120
      Top             =   3120
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdComixCxEXT 
      BackColor       =   &H8000000D&
      Caption         =   "Comiss. CX EXTRA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaskColor       =   &H00FF0000&
      TabIndex        =   118
      Top             =   2760
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdComisCx 
      BackColor       =   &H8000000D&
      Caption         =   "Comiss. CX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaskColor       =   &H00FF0000&
      TabIndex        =   117
      Top             =   2400
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdComisMwts 
      BackColor       =   &H8000000D&
      Caption         =   "Comiss. Vendedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaskColor       =   &H00FF0000&
      TabIndex        =   116
      Top             =   2040
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.Frame frmQtde 
      BackColor       =   &H0080FF80&
      Caption         =   " QTDES "
      Height          =   1815
      Left            =   11040
      TabIndex        =   122
      Top             =   6720
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Frame Frame5 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         TabIndex        =   138
         Top             =   1320
         Width           =   1815
         Begin VB.TextBox txtQtdeLimiteV 
            Alignment       =   2  'Center
            Enabled         =   0   'False
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
            Left            =   1200
            TabIndex        =   139
            Top             =   50
            Width           =   495
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "LIM. VEND.:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   140
            Top             =   60
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2160
         TabIndex        =   135
         Top             =   1320
         Width           =   1815
         Begin VB.TextBox txtQtdeTOTAL 
            Alignment       =   2  'Center
            Enabled         =   0   'False
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
            Left            =   960
            TabIndex        =   137
            Top             =   50
            Width           =   735
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   136
            Top             =   45
            Width           =   855
         End
      End
      Begin VB.TextBox txtQtdeX 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         TabIndex        =   134
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtQtdeC 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   132
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtQtdeG 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   130
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtQtdeDEM 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         TabIndex        =   128
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtQtdeR 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         TabIndex        =   126
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtQtdeV 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   124
         Top             =   960
         Width           =   495
      End
      Begin Skin_Button.ctr_Button cmdRelQtdeTipo 
         Height          =   285
         Left            =   3800
         TabIndex        =   141
         TabStop         =   0   'False
         ToolTipText     =   "Relatório de Funcionários por Função"
         Top             =   120
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   503
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   12632319
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_Alt_Fic_Mensal_VIS.frx":264A
         PICN            =   "frm_Alt_Fic_Mensal_VIS.frx":2666
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "CAIXA EXT.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   133
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "CAIXA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   131
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "GERENTE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "DEMITIDOS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   127
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "SEGURANÇA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   125
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "VENDEDOR:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   123
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRelAdmin 
      Caption         =   "*"
      Height          =   255
      Left            =   10560
      TabIndex        =   119
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdAtualizarCaption 
      Appearance      =   0  'Flat
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
      Height          =   795
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":3948
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   900
      Width           =   225
   End
   Begin VB.Frame frmAtualizando 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   2520
      TabIndex        =   111
      Top             =   5160
      Visible         =   0   'False
      Width           =   5640
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ATUALIZANDO..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   705
         Left            =   240
         TabIndex        =   112
         Top             =   120
         Width           =   5160
      End
   End
   Begin VB.CommandButton cmdAtualizar 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
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
      Left            =   210
      Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":39BF
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   960
      Width           =   465
   End
   Begin MSDataGridLib.DataGrid Grid 
      Bindings        =   "frm_Alt_Fic_Mensal_VIS.frx":3CC9
      CausesValidation=   0   'False
      Height          =   6915
      Left            =   10920
      TabIndex        =   0
      Top             =   1680
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   12197
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   12648447
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
      ColumnCount     =   43
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
         DataField       =   "M_NFICHA"
         Caption         =   "M_NFICHA"
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
         DataField       =   "M_F_COD"
         Caption         =   "M_F_COD"
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
         DataField       =   "M_LOGO"
         Caption         =   "M_LOGO"
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
         DataField       =   "M_FERIAS"
         Caption         =   "M_FERIAS"
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
      BeginProperty Column07 
         DataField       =   "M_OBS"
         Caption         =   "M_OBS"
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
         DataField       =   "M_ANOTACAO"
         Caption         =   "M_ANOTACAO"
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
      BeginProperty Column09 
         DataField       =   "M_DT_ADM"
         Caption         =   "M_DT_ADM"
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
         DataField       =   "M_DT_REG"
         Caption         =   "M_DT_REG"
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
         DataField       =   "M_DT_DEM"
         Caption         =   "M_DT_DEM"
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
         DataField       =   "M_DEM_OK"
         Caption         =   "M_DEM_OK"
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
         DataField       =   "M_BLOQ"
         Caption         =   "M_BLOQ"
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
         DataField       =   "M_TOTAL_VND"
         Caption         =   "M_TOTAL_VND"
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
         DataField       =   "M_TOTAL_COM"
         Caption         =   "M_TOTAL_COM"
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
         DataField       =   "M_TOTAL_PRE"
         Caption         =   "M_TOTAL_PRE"
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
      BeginProperty Column17 
         DataField       =   "M_TOTAL"
         Caption         =   "M_TOTAL"
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
      BeginProperty Column18 
         DataField       =   "M_TOTAL_MAIS"
         Caption         =   "M_TOTAL_MAIS"
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
      BeginProperty Column19 
         DataField       =   "M_TOTAL_MENOS"
         Caption         =   "M_TOTAL_MENOS"
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
      BeginProperty Column20 
         DataField       =   "M_NOME"
         Caption         =   "FUNCIONARIO"
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
      BeginProperty Column21 
         DataField       =   "M_FERIAS_PG"
         Caption         =   "M_FERIAS_PG"
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
      BeginProperty Column22 
         DataField       =   "M_FERIAS_Ult_PG"
         Caption         =   "M_FERIAS_Ult_PG"
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
      BeginProperty Column23 
         DataField       =   "M_FERIAS_OK"
         Caption         =   "M_FERIAS_OK"
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
      BeginProperty Column24 
         DataField       =   "M_13_PG"
         Caption         =   "M_13_PG"
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
      BeginProperty Column25 
         DataField       =   "M_13_ULT_PG"
         Caption         =   "M_13_ULT_PG"
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
      BeginProperty Column26 
         DataField       =   "M_13_OBS"
         Caption         =   "M_13_OBS"
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
      BeginProperty Column27 
         DataField       =   "M_13_OK"
         Caption         =   "M_13_OK"
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
      BeginProperty Column28 
         DataField       =   "M_DT_ACF"
         Caption         =   "M_DT_ACF"
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
      BeginProperty Column29 
         DataField       =   "M_EMPRESTIMO_ANOT"
         Caption         =   "M_EMPRESTIMO_ANOT"
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
      BeginProperty Column30 
         DataField       =   "M_ACORDO"
         Caption         =   "M_ACORDO"
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
      BeginProperty Column31 
         DataField       =   "M_VCTO_FERIAS"
         Caption         =   "M_VCTO_FERIAS"
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
      BeginProperty Column32 
         DataField       =   "M_COMISSAO"
         Caption         =   "M_COMISSAO"
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
      BeginProperty Column33 
         DataField       =   "M_Vnd"
         Caption         =   "M_Vnd"
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
      BeginProperty Column34 
         DataField       =   "M_NUM_FILHOS"
         Caption         =   "M_NUM_FILHOS"
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
      BeginProperty Column35 
         DataField       =   "M_PG_SAL_FAM"
         Caption         =   "M_PG_SAL_FAM"
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
      BeginProperty Column36 
         DataField       =   "M_NOTAS"
         Caption         =   "M_NOTAS"
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
      BeginProperty Column37 
         DataField       =   "BLOQ"
         Caption         =   "#"
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
      BeginProperty Column38 
         DataField       =   "F_LOJA"
         Caption         =   "F_LOJA"
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
      BeginProperty Column39 
         DataField       =   "F_TIPO"
         Caption         =   "T"
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
      BeginProperty Column40 
         DataField       =   "F_COD_CENTRAL"
         Caption         =   "F_COD_CENTRAL"
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
      BeginProperty Column41 
         DataField       =   "F_DEM_OK"
         Caption         =   "F_DEM_OK"
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
      BeginProperty Column42 
         DataField       =   "num"
         Caption         =   "num"
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
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column16 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column17 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column18 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column19 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   2670,236
         EndProperty
         BeginProperty Column21 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column22 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column23 
            Object.Visible         =   0   'False
            ColumnWidth     =   1184,882
         EndProperty
         BeginProperty Column24 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column25 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column26 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column27 
            Object.Visible         =   0   'False
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column28 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column29 
            Object.Visible         =   0   'False
            ColumnWidth     =   1920,189
         EndProperty
         BeginProperty Column30 
            Object.Visible         =   0   'False
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column31 
            Object.Visible         =   0   'False
            ColumnWidth     =   1395,213
         EndProperty
         BeginProperty Column32 
            Object.Visible         =   0   'False
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column33 
            Object.Visible         =   0   'False
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column34 
            Object.Visible         =   0   'False
            ColumnWidth     =   1335,118
         EndProperty
         BeginProperty Column35 
            Object.Visible         =   0   'False
            ColumnWidth     =   1335,118
         EndProperty
         BeginProperty Column36 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column37 
            ColumnWidth     =   345,26
         EndProperty
         BeginProperty Column38 
            Object.Visible         =   0   'False
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column39 
            ColumnWidth     =   299,906
         EndProperty
         BeginProperty Column40 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column41 
            Object.Visible         =   0   'False
            ColumnWidth     =   929,764
         EndProperty
         BeginProperty Column42 
            Object.Visible         =   0   'False
            ColumnWidth     =   540,284
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmBloq 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   8040
      TabIndex        =   106
      Top             =   6240
      Visible         =   0   'False
      Width           =   2760
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BLOQUEADO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   107
         Top             =   0
         Width           =   2760
      End
   End
   Begin VB.CommandButton cmdDesbloquear 
      Caption         =   "Bloq/Liberar esta ficha"
      Height          =   255
      Left            =   5160
      TabIndex        =   105
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin Skin_Button.ctr_Button cmdAddSalario 
      Height          =   285
      Left            =   10200
      TabIndex        =   89
      TabStop         =   0   'False
      ToolTipText     =   "Gerar COMISSÃO para este funcionário"
      Top             =   6960
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
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
      MICON           =   "frm_Alt_Fic_Mensal_VIS.frx":3CDE
      PICN            =   "frm_Alt_Fic_Mensal_VIS.frx":3CFA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000000&
      Caption         =   " Mostrar: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   13680
      TabIndex        =   101
      Top             =   840
      Width           =   1695
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "IR"
         Height          =   255
         Left            =   1080
         TabIndex        =   103
         Top             =   560
         Width           =   495
      End
      Begin VB.ComboBox cbMostrar 
         Height          =   315
         ItemData        =   "frm_Alt_Fic_Mensal_VIS.frx":591C
         Left            =   120
         List            =   "frm_Alt_Fic_Mensal_VIS.frx":5941
         TabIndex        =   102
         Text            =   "TODOS"
         Top             =   225
         Width           =   1515
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   99
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frmOrdem 
      BackColor       =   &H80000000&
      Caption         =   " Ordem: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   12000
      TabIndex        =   96
      Top             =   840
      Width           =   1575
      Begin VB.CommandButton cmdEsconder 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1170
         MaskColor       =   &H8000000A&
         Style           =   1  'Graphical
         TabIndex        =   144
         Top             =   120
         Width           =   390
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Tipo"
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
         Left            =   720
         TabIndex        =   109
         Top             =   480
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optLoja 
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
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optNome 
         Caption         =   "Nome"
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
         Left            =   120
         TabIndex        =   97
         Top             =   240
         Width           =   855
      End
   End
   Begin Skin_Button.ctr_Button cmdFixosSaldos 
      Height          =   285
      Left            =   6840
      TabIndex        =   94
      TabStop         =   0   'False
      ToolTipText     =   "Lançamentos Fixos e Saldos Negativos"
      Top             =   1320
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   12632319
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_Alt_Fic_Mensal_VIS.frx":59AC
      PICN            =   "frm_Alt_Fic_Mensal_VIS.frx":59C8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Skin_Button.ctr_Button cmdDescCalcFixo 
      Height          =   285
      Left            =   9840
      TabIndex        =   93
      TabStop         =   0   'False
      ToolTipText     =   "Gerar Lançamentos Fixos"
      Top             =   6960
      Visible         =   0   'False
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
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
      MICON           =   "frm_Alt_Fic_Mensal_VIS.frx":6CAA
      PICN            =   "frm_Alt_Fic_Mensal_VIS.frx":6CC6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txt_notas 
      DataField       =   "M_NOTAS"
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
      Height          =   1125
      Left            =   1800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   91
      Top             =   5880
      Visible         =   0   'False
      Width           =   8970
   End
   Begin VB.Frame CARIMBO 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   1545
      Left            =   9240
      TabIndex        =   63
      Top             =   3000
      Visible         =   0   'False
      Width           =   1470
      Begin VB.Label lb_Dt_ACF 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "20/02/2006"
         DataField       =   "M_DT_ACF"
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
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   195
         TabIndex        =   65
         Top             =   975
         Width           =   1050
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ACERTO FINAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   705
         Left            =   -15
         TabIndex        =   64
         Top             =   405
         Width           =   1560
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000C000&
         BorderWidth     =   7
         X1              =   600
         X2              =   390
         Y1              =   1200
         Y2              =   600
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000C000&
         BorderWidth     =   7
         X1              =   510
         X2              =   1035
         Y1              =   1200
         Y2              =   435
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0FFC0&
         BorderColor     =   &H000000C0&
         BorderWidth     =   7
         FillColor       =   &H000000FF&
         Height          =   1335
         Left            =   -720
         Shape           =   3  'Circle
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.CheckBox ck_pg_SFam 
      Caption         =   "Check1"
      DataField       =   "M_PG_SAL_FAM"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      Height          =   255
      Left            =   9480
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   4830
      Width           =   195
   End
   Begin VB.TextBox txt_NFilhos 
      Alignment       =   2  'Center
      DataField       =   "M_NUM_FILHOS"
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
      Height          =   315
      Left            =   8445
      MaxLength       =   10
      TabIndex        =   82
      Top             =   4845
      Width           =   855
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
      Left            =   8430
      MaxLength       =   10
      TabIndex        =   79
      Top             =   5310
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
      Left            =   9600
      MaxLength       =   10
      TabIndex        =   78
      Top             =   5280
      Width           =   1095
   End
   Begin VB.ComboBox txt_Vcto_ferias 
      DataField       =   "M_VCTO_FERIAS"
      DataSource      =   "adoReg"
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frm_Alt_Fic_Mensal_VIS.frx":88E8
      Left            =   120
      List            =   "frm_Alt_Fic_Mensal_VIS.frx":8910
      TabIndex        =   11
      Top             =   3900
      Width           =   660
   End
   Begin VB.CheckBox ck_Acordo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Acordo?"
      DataField       =   "M_Acordo"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      Height          =   225
      Left            =   6480
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   2160
      Width           =   915
   End
   Begin Skin_Button.ctr_Button btRptDem 
      Height          =   285
      Left            =   3090
      TabIndex        =   68
      TabStop         =   0   'False
      ToolTipText     =   "Relatório dos (D)"
      Top             =   3060
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   503
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   12632319
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_Alt_Fic_Mensal_VIS.frx":893B
      PICN            =   "frm_Alt_Fic_Mensal_VIS.frx":8957
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton cmdATTotal 
      Caption         =   "A.T"
      Enabled         =   0   'False
      Height          =   315
      Left            =   12600
      Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":9C39
      Style           =   1  'Graphical
      TabIndex        =   66
      TabStop         =   0   'False
      ToolTipText     =   "Atualiza os Totais de todas as Fichas!"
      Top             =   8640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   11160
      Top             =   2640
   End
   Begin VB.TextBox TXT_AC_F 
      DataField       =   "M_DT_ACF"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      DataSource      =   "ADOREG"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   630
      Left            =   5160
      Locked          =   -1  'True
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   62
      Top             =   2400
      Width           =   4380
   End
   Begin VB.CheckBox CK_ACF 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ac. FINAL"
      DataField       =   "M_BLOQ"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   5160
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1260
   End
   Begin VB.CheckBox CK_DEM 
      Caption         =   "Check1"
      DataField       =   "M_DEM_OK"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2670
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   3075
      Width           =   195
   End
   Begin VB.CheckBox CK_FERIAS 
      Caption         =   "Check1"
      DataField       =   "M_FERIAS_OK"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3480
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   3960
      Width           =   195
   End
   Begin VB.CheckBox CK_13 
      Caption         =   "Check1"
      DataField       =   "M_13_OK"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10485
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   3390
      Width           =   195
   End
   Begin VB.TextBox TXT_13_OBS 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      DataField       =   "M_13_OBS"
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
      Left            =   4680
      TabIndex        =   9
      Top             =   3315
      Width           =   4695
   End
   Begin rdActiveText.ActiveText TXT_FERIAS_PG 
      DataField       =   "M_FERIAS_PG"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      DataSource      =   "ADOREG"
      Height          =   315
      Left            =   2400
      TabIndex        =   13
      Top             =   3900
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Enabled         =   0   'False
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
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin MSDataListLib.DataCombo TXT_CRED 
      Bindings        =   "frm_Alt_Fic_Mensal_VIS.frx":9F43
      DataField       =   "M_F_COD"
      DataSource      =   "ADOREG"
      Height          =   315
      Left            =   240
      TabIndex        =   24
      Top             =   7650
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
   Begin VB.TextBox TXT_MAIS 
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   11325
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   8655
      Width           =   1350
   End
   Begin VB.TextBox TXT_MENOS 
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
      ForeColor       =   &H008080FF&
      Height          =   285
      Left            =   11325
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   8940
      Width           =   1350
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
      Left            =   14160
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   8745
      Width           =   1170
   End
   Begin VB.TextBox txt_F_COD 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
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
      Left            =   4365
      TabIndex        =   33
      Top             =   2640
      Width           =   660
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
      Left            =   14160
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   9165
      Width           =   1170
   End
   Begin VB.TextBox TXT_OBS 
      DataField       =   "M_OBS"
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
      Height          =   645
      Left            =   75
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   4920
      Width           =   8250
   End
   Begin VB.TextBox TXT_TOTAL 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   11325
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   9240
      Width           =   1350
   End
   Begin MSAdodcLib.Adodc ADO_LANC 
      Height          =   330
      Left            =   3000
      Top             =   9210
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
      BackColor       =   8454143
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
   Begin VB.TextBox TXT_ANO 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
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
      Left            =   4440
      TabIndex        =   3
      Top             =   2085
      Width           =   570
   End
   Begin VB.ComboBox TXT_MES 
      BackColor       =   &H00C0FFC0&
      DataField       =   "M_MES"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frm_Alt_Fic_Mensal_VIS.frx":9F54
      Left            =   3720
      List            =   "frm_Alt_Fic_Mensal_VIS.frx":9F7C
      TabIndex        =   2
      Top             =   2085
      Width           =   660
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
      Left            =   9720
      TabIndex        =   17
      Top             =   2640
      Width           =   930
   End
   Begin VB.TextBox TXT_FERIAS 
      DataField       =   "M_FERIAS"
      DataSource      =   "ADOREG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   75
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   4215
      Width           =   5175
   End
   Begin VB.TextBox TXT_ANOTACAO 
      DataField       =   "M_ANOTACAO"
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
      Height          =   675
      Left            =   5280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   3975
      Width           =   5385
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   14760
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
            Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":9FA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":A2C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":A5DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":A8F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":AC0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":AF29
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":B243
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":BB1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":D827
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":DB41
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":DF93
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":E2AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":E5CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":10D81
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":111D3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo TXT_FUNC 
      Bindings        =   "frm_Alt_Fic_Mensal_VIS.frx":13985
      DataField       =   "M_F_COD"
      DataSource      =   "ADOREG"
      Height          =   315
      Left            =   75
      TabIndex        =   16
      Top             =   2640
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      Style           =   2
      BackColor       =   12648384
      ListField       =   "F_NOME"
      BoundColumn     =   "F_Codigo"
      Text            =   ""
      Object.DataMember      =   "TAB_FUNCIONARIO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo txtLogo 
      Bindings        =   "frm_Alt_Fic_Mensal_VIS.frx":13996
      DataField       =   "M_LOGO"
      DataSource      =   "ADOREG"
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2085
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      BackColor       =   12648384
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
   Begin rdActiveText.ActiveText TXT_FERIAS_ULT_PG 
      DataField       =   "M_FERIAS_ULT_PG"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      DataSource      =   "ADOREG"
      Height          =   315
      Left            =   1140
      TabIndex        =   12
      Top             =   3900
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Enabled         =   0   'False
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
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText TXT_13_PG 
      DataField       =   "M_13_PG"
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
      Left            =   9480
      TabIndex        =   10
      Top             =   3300
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      Enabled         =   0   'False
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
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText TXT_13_ULT_PG 
      DataField       =   "M_13_ULT_PG"
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
      Height          =   330
      Left            =   3675
      TabIndex        =   8
      Top             =   3315
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      Enabled         =   0   'False
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
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txt_DT_ADM 
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
      Height          =   315
      Left            =   105
      TabIndex        =   5
      Top             =   3315
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Enabled         =   0   'False
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
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText TXT_DT_REG 
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
      Height          =   315
      Left            =   1230
      TabIndex        =   6
      Top             =   3315
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Enabled         =   0   'False
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
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText TXT_DT_DEM 
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
      Height          =   315
      Left            =   2355
      TabIndex        =   7
      Top             =   3315
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Enabled         =   0   'False
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
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin Skin_Button.ctr_Button btRptVctoFerias 
      Height          =   285
      Left            =   765
      TabIndex        =   74
      TabStop         =   0   'False
      ToolTipText     =   "Relatório das Férias Vencendo"
      Top             =   3915
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   503
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   12632319
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_Alt_Fic_Mensal_VIS.frx":139A7
      PICN            =   "frm_Alt_Fic_Mensal_VIS.frx":139C3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc ADOREG 
      Height          =   330
      Left            =   0
      Top             =   9600
      Visible         =   0   'False
      Width           =   15480
      _ExtentX        =   27305
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
      BackColor       =   8454143
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
   Begin Skin_Button.ctr_Button btRptADM 
      Height          =   285
      Left            =   510
      TabIndex        =   75
      TabStop         =   0   'False
      ToolTipText     =   "Relatório dos @"
      Top             =   3060
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   503
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   12632319
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_Alt_Fic_Mensal_VIS.frx":14CA5
      PICN            =   "frm_Alt_Fic_Mensal_VIS.frx":14CC1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Skin_Button.ctr_Button cmdAddLanç_SalF 
      Height          =   525
      Left            =   10200
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   4560
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
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
      MICON           =   "frm_Alt_Fic_Mensal_VIS.frx":15FA3
      PICN            =   "frm_Alt_Fic_Mensal_VIS.frx":15FBF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Skin_Button.ctr_Button cmdSalarioGerente 
      Height          =   285
      Left            =   5760
      TabIndex        =   90
      TabStop         =   0   'False
      ToolTipText     =   "Salários Gerentes"
      Top             =   1320
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   503
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   12632319
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_Alt_Fic_Mensal_VIS.frx":17BE1
      PICN            =   "frm_Alt_Fic_Mensal_VIS.frx":17BFD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
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
      Height          =   840
      Left            =   720
      TabIndex        =   38
      Top             =   840
      Width           =   11175
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
         Height          =   585
         Left            =   6960
         TabIndex        =   41
         Top             =   180
         Visible         =   0   'False
         Width           =   1815
         Begin VB.ComboBox txt_PMes 
            Height          =   315
            ItemData        =   "frm_Alt_Fic_Mensal_VIS.frx":18EDF
            Left            =   195
            List            =   "frm_Alt_Fic_Mensal_VIS.frx":18F07
            TabIndex        =   42
            Top             =   195
            Width           =   570
         End
         Begin VB.TextBox txt_PAno 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   960
            TabIndex        =   43
            Top             =   195
            Width           =   615
         End
         Begin VB.Line Line1 
            BorderWidth     =   3
            X1              =   720
            X2              =   900
            Y1              =   465
            Y2              =   240
         End
      End
      Begin VB.TextBox TXT_AC_F_Modelo 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   435
         Left            =   6720
         MaxLength       =   58
         MultiLine       =   -1  'True
         TabIndex        =   76
         Top             =   240
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.OptionButton Op 
         Caption         =   "VCTO (F)"
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
         Index           =   7
         Left            =   2550
         TabIndex        =   73
         Top             =   555
         Width           =   1095
      End
      Begin VB.OptionButton Op 
         Caption         =   "(D)"
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
         Index           =   6
         Left            =   1575
         TabIndex        =   67
         ToolTipText     =   "Todos com Empréstimo"
         Top             =   555
         Width           =   900
      End
      Begin VB.OptionButton Op 
         Caption         =   "S. Emprést."
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
         Index           =   4
         Left            =   120
         TabIndex        =   54
         ToolTipText     =   "Todos com Empréstimo"
         Top             =   555
         Width           =   1320
      End
      Begin VB.OptionButton Op 
         Caption         =   "Nº Ficha"
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
         Index           =   0
         Left            =   135
         TabIndex        =   48
         Top             =   255
         Width           =   960
      End
      Begin VB.OptionButton Op 
         Caption         =   "(B)"
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
         Index           =   1
         Left            =   1590
         TabIndex        =   47
         Top             =   255
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Op 
         Caption         =   "Mês / Ano"
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
         Index           =   2
         Left            =   2565
         TabIndex        =   46
         Top             =   270
         Width           =   1200
      End
      Begin VB.OptionButton Op 
         Caption         =   "Nome"
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
         Index           =   3
         Left            =   3990
         TabIndex        =   45
         Top             =   555
         Width           =   855
      End
      Begin VB.OptionButton Op 
         Caption         =   "Remover Filtro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   5
         Left            =   4800
         TabIndex        =   44
         Top             =   255
         Width           =   1575
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
         Left            =   6960
         TabIndex        =   39
         Top             =   195
         Visible         =   0   'False
         Width           =   3495
         Begin VB.TextBox txt_Pesq 
            Height          =   285
            Left            =   240
            TabIndex        =   40
            Top             =   210
            Width           =   3015
         End
      End
      Begin VB.CommandButton cmdFiltrar 
         Height          =   555
         Left            =   10500
         Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":18F32
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   195
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CommandButton bt_Salva_Ac 
         Height          =   555
         Left            =   10500
         Picture         =   "frm_Alt_Fic_Mensal_VIS.frx":1923C
         Style           =   1  'Graphical
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   195
         Width           =   600
      End
      Begin Skin_Button.ctr_Button bt_VoltarDT 
         Height          =   315
         Left            =   3840
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   195
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
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
         MICON           =   "frm_Alt_Fic_Mensal_VIS.frx":19546
         PICN            =   "frm_Alt_Fic_Mensal_VIS.frx":19562
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button bt_AvaDT 
         Height          =   315
         Left            =   4200
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   195
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
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
         MICON           =   "frm_Alt_Fic_Mensal_VIS.frx":199B4
         PICN            =   "frm_Alt_Fic_Mensal_VIS.frx":199D0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button cmdComissaoVendedor 
         Height          =   285
         Left            =   5760
         TabIndex        =   150
         TabStop         =   0   'False
         ToolTipText     =   "Comissão Vendedores"
         Top             =   480
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   12632319
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_Alt_Fic_Mensal_VIS.frx":19E22
         PICN            =   "frm_Alt_Fic_Mensal_VIS.frx":19E3E
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
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   100
      Top             =   0
      Width           =   15450
      _ExtentX        =   27252
      _ExtentY        =   1482
      ButtonWidth     =   1693
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
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
            Caption         =   "&Novo"
            Key             =   "nova"
            Object.ToolTipText     =   "Nova Ficha"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Editar"
            Key             =   "editar"
            Object.ToolTipText     =   "Editar Alteração (Alt+E)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Salvar"
            Key             =   "salvar"
            Object.ToolTipText     =   "Salvar Alteração (Alt+S)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
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
            Object.ToolTipText     =   "Excluir FICHA (Alt+X)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Key             =   "imprimir"
            Object.ToolTipText     =   "Imprime as Fichas"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "C&ontas"
            Key             =   "conta"
            Object.ToolTipText     =   "Incluir Contas (Alt+O)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "(F5)"
            Key             =   "dupla"
            Object.ToolTipText     =   "Visualizar Ficha Dupla"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Vistar Ct."
            Key             =   "vistar"
            Object.ToolTipText     =   "Vistar Contas"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Bloq/Lib"
            Key             =   "desbloquear"
            Object.ToolTipText     =   "Desbloquear Fichas"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Comissão"
            Key             =   "gcomissao"
            Object.ToolTipText     =   "Gerar Comissões"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Emprést."
            Key             =   "emp"
            Object.ToolTipText     =   "Consultar Empréstimo"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Prog."
            Key             =   "programados"
            Description     =   "Lançamentos Programados"
            Object.ToolTipText     =   "Lançamentos Programados"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ca&dastro"
            Key             =   "cadastro"
            Description     =   "Ficha Cadastral"
            Object.ToolTipText     =   "Ficha Cadastral"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ger / CX"
            Key             =   "gerente"
            Description     =   "Comissão Gerentes / Caixas"
            Object.ToolTipText     =   "Comissão Gerentes / Caixas"
            ImageIndex      =   15
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexGRID_L 
      Bindings        =   "frm_Alt_Fic_Mensal_VIS.frx":1B120
      Height          =   2535
      Left            =   0
      TabIndex        =   95
      Top             =   7080
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   10
      FixedRows       =   0
      FixedCols       =   0
      ForeColorSel    =   -2147483639
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      ScrollBars      =   2
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
   Begin MSDataListLib.DataCombo txtLogo2 
      Bindings        =   "frm_Alt_Fic_Mensal_VIS.frx":1B137
      DataField       =   "M_LOGO"
      DataSource      =   "ADOREG"
      Height          =   495
      Left            =   1920
      TabIndex        =   142
      TabStop         =   0   'False
      Top             =   1920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      BackColor       =   12648384
      ListField       =   "NUM"
      BoundColumn     =   "COD_LOJ"
      Text            =   ""
      Object.DataMember      =   "tab_L_num"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Skin_Button.ctr_Button btRptREG 
      Height          =   285
      Left            =   1560
      TabIndex        =   157
      TabStop         =   0   'False
      ToolTipText     =   "Relatório dos ®"
      Top             =   3060
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   503
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   12632319
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_Alt_Fic_Mensal_VIS.frx":1B148
      PICN            =   "frm_Alt_Fic_Mensal_VIS.frx":1B164
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label txtTempoServico 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8 anos"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "mmmm-yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   161
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label txtIdade 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "25 anos"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "mmmm-yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   160
      Top             =   5835
      Width           =   1095
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "TEMPO DE SERVIÇO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   159
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "IDADE:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   158
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Pg. VT"
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
      Left            =   3720
      TabIndex        =   148
      Top             =   3840
      Width           =   705
   End
   Begin VB.Label lblMes 
      Alignment       =   2  'Center
      Caption         =   "MÊS - ANO"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "mmmm-yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   7680
      TabIndex        =   114
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label TXT_FTIPO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FUNCAO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   104
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblNotas 
      BackStyle       =   0  'Transparent
      Caption         =   "ANOTAÇÕES EXTRAS:"
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
      Left            =   1800
      TabIndex        =   92
      Top             =   5640
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Pg. Sal. F."
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
      Left            =   9390
      TabIndex        =   84
      Top             =   4650
      Width           =   840
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
      Left            =   8460
      TabIndex        =   83
      Top             =   4650
      Width           =   960
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "V. PISO BRT"
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
      Left            =   8445
      TabIndex        =   81
      Top             =   5115
      Width           =   960
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "V. PISO LIQ."
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
      Left            =   9615
      TabIndex        =   80
      Top             =   5115
      Width           =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(F) ULT:"
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
      Left            =   1110
      TabIndex        =   72
      Top             =   3630
      Width           =   1110
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "(F) PG:"
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
      Left            =   2385
      TabIndex        =   71
      Top             =   3630
      Width           =   735
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Vcto (F):"
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
      Left            =   150
      TabIndex        =   70
      Top             =   3630
      Width           =   855
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "OBS 13º:"
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
      Left            =   4680
      TabIndex        =   57
      Top             =   3090
      Width           =   855
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "(13º) ULT:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   56
      Top             =   3105
      Width           =   855
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "(13º) PG:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9405
      TabIndex        =   55
      Top             =   3090
      Width           =   735
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
      Left            =   10995
      TabIndex        =   53
      Top             =   8550
      Width           =   240
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
      Left            =   11040
      TabIndex        =   52
      Top             =   8805
      Width           =   150
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "(D)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2385
      TabIndex        =   37
      Top             =   3090
      Width           =   540
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
      Left            =   12825
      TabIndex        =   36
      Top             =   8790
      Width           =   1365
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº:"
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
      Left            =   4365
      TabIndex        =   34
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "®"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1245
      TabIndex        =   32
      Top             =   3090
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
      Left            =   150
      TabIndex        =   31
      Top             =   3075
      Width           =   1575
   End
   Begin VB.Label lbl_SaldoAnt 
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
      Left            =   12915
      TabIndex        =   30
      Top             =   9180
      Visible         =   0   'False
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
      Left            =   2880
      TabIndex        =   28
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVAÇÃO:"
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
      Left            =   75
      TabIndex        =   27
      Top             =   4695
      Width           =   2175
   End
   Begin VB.Label lbl_total 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "T."
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
      Left            =   11040
      TabIndex        =   26
      Top             =   9240
      Visible         =   0   'False
      Width           =   225
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
      TabIndex        =   23
      Top             =   7530
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº FICHA:"
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
      Left            =   9735
      TabIndex        =   22
      Top             =   2400
      Width           =   1080
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
      Left            =   3720
      TabIndex        =   21
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ANOTAÇÃO:"
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
      Left            =   5280
      TabIndex        =   20
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NOME:"
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
      Left            =   75
      TabIndex        =   19
      Top             =   2415
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
      Left            =   4440
      TabIndex        =   18
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   5400
      Left            =   0
      Top             =   1680
      Width           =   10890
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      Height          =   915
      Left            =   12825
      Top             =   8640
      Width           =   2550
   End
   Begin VB.Menu mnu 
      Caption         =   "Menu"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuVis 
         Caption         =   "Vistar"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuRem 
         Caption         =   "Remover"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVisT 
         Caption         =   "Vistar Todos"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuRemT 
         Caption         =   "Remover Todos"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuAcessoTotal 
         Caption         =   "Acesso Total"
         Shortcut        =   {F4}
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frm_Alt_Fic_Mensal_VIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W_LD_FILTRO As Boolean
Dim V_MOVE As Boolean
Dim V_MOVE_GRID As Boolean
Dim W_FILTRO As String
Dim W_INDEX As Byte
Dim w_F5 As Boolean
Dim W_CK_DEM As Boolean
Dim w_SN_Total As Boolean
Dim w_reset_tipo As Boolean
Dim wTxtOld As String
Dim w_logo As String
Dim w_ck_vt As Boolean
Dim w_ac_f As String

'--------- flex grid -------------------------------------
Private ControlVisible As Boolean     ' Se o controle esta visivel ou nao
Private LastRow As Long               ' Ultima linha em que se editou
Private LastCol As Long               ' ultima coluna em que se editou


Sub Lancamentos()
    
    If (ADOREG.Recordset.Fields("M_TIPO") = "V" Or ADOREG.Recordset.Fields("M_TIPO") = "C" Or ADOREG.Recordset.Fields("M_TIPO") = "X" Or ADOREG.Recordset.Fields("M_TIPO") = "2") Or acessoTotal() Then
        If de.rscmdSqlVisAltContas3.State = 1 Then de.rscmdSqlVisAltContas3.Close
        de.cmdSqlVisAltContas3 ADOREG.Recordset.Fields("M_NFICHA")
        Set ADO_LANC.Recordset = de.rscmdSqlVisAltContas3.Clone
    Else
        If de.rscmdSqlVisAltContas2.State = 1 Then de.rscmdSqlVisAltContas2.Close
        de.cmdSqlVisAltContas2 ADOREG.Recordset.Fields("M_NFICHA")
        Set ADO_LANC.Recordset = de.rscmdSqlVisAltContas2.Clone
    End If

    flexGRID_L.ColAlignment(4) = flexAlignRightBottom 'valor
    flexGRID_L.ColAlignment(5) = flexAlignCenterBottom 'op

   'Varrendo todas as linhas do flexgrid
   For I = 1 To flexGRID_L.Rows - 1
   If flexGRID_L.TextMatrix(I, 4) < 0 Then
        For coluna = 0 To flexGRID_L.Cols - 1
            flexGRID_L.Col = coluna
            flexGRID_L.Row = I
            flexGRID_L.CellFontBold = True
            flexGRID_L.CellForeColor = vbRed
        Next coluna
   ElseIf flexGRID_L.TextMatrix(I, 4) > 0 Then
        For coluna = 0 To flexGRID_L.Cols - 1
            flexGRID_L.Col = coluna
            flexGRID_L.Row = I
            flexGRID_L.CellFontBold = True
            flexGRID_L.CellForeColor = vbBlue
        Next coluna
   End If
   If flexGRID_L.TextMatrix(I, 8) > 0 Then
        ' Varre todas as colunas da linha e seta a cor de fundo
        For coluna = 0 To flexGRID_L.Cols - 1
            flexGRID_L.Col = coluna
            flexGRID_L.Row = I
            flexGRID_L.CellBackColor = vbYellow
            flexGRID_L.CellFontBold = True
            'flexGRID_L.CellForeColor = &H80000004
        Next coluna
    End If
    'If Len(flexGRID_L.TextMatrix(I, 2)) > 75 Then
            'flexGRID_L.Col = 2
            'flexGRID_L.Row = I
            'flexGRID_L.CellFontSize = 7
    'End If
Next I

        'formatarFlexGrid
        
        
        'ATUALIZA DATA DE DEMISSÃO DO CADASTRO
        w_dt_DEM = IIf(TXT_DT_DEM = "", Null, Format(TXT_DT_DEM, "DD/MM/YYYY"))
        'Se não tiver fichas abertas
        If de.cnc.Execute("SELECT M_NFICHA FROM TAB_FICHA_MENS WHERE M_DT_DEM IS NULL AND M_F_COD = " & txt_F_COD & " AND ((M_MES >= " & TXT_MES & " AND M_ANO = " & TXT_ANO & ") OR M_ANO > " & TXT_ANO & ")").RecordCount = 0 Then
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_DT_DEM = '" & w_dt_DEM & "'  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        Else
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_DT_DEM = NULL  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        End If


End Sub

Private Sub ADO_LANC_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If w_SN_Total = True Then Total
End Sub

Private Sub bt_Salva_Ac_Click()
On Error Resume Next
    
    If vbYes = MsgBox("Deseja inserir Ac. Final p/ todos os registros demitidos?", vbQuestion + vbYesNo) Then
    w_SN_Total = False
            ADOREG.Recordset.MoveFirst
            Do While Not ADOREG.Recordset.EOF
                ADOREG.Recordset.Fields("M_DT_ACF") = TXT_AC_F_Modelo
                ADOREG.Recordset.Fields("M_BLOQ") = True
                ADOREG.Recordset.Fields("M_DEM_OK") = True
                'ADOREG.Recordset.UpdateBatch adAffectCurrent
                ADOREG.Recordset.MoveNext
            Loop
            ADOREG.Recordset.MoveFirst
    w_SN_Total = True
    End If
End Sub

Sub bt_AvaDT_Click()
On Error Resume Next
    Dim w_cod_atual As String
    
    w_cod_atual = txt_F_COD
    
    w_reset_tipo = False

        If CDbl(txt_PMes) = 12 Then
        txt_PMes = 1
        txt_PAno = CDbl(txt_PAno) + 1
    Else
        txt_PMes = Format(CDbl(txt_PMes) + 1, "00")
    End If
    
    wID = W_INDEX
    W_INDEX = 2
    Call cmdFiltrar_Click
    
    W_INDEX = wID
    Call cmdFiltrar_Click
    
    Op(2).Caption = Format(CDbl(txt_PMes), "00") & "/" & txt_PAno
    
    If Not (ADOREG.Recordset.EOF) Then
    
        If optLoja.value Then ADOREG.Recordset.Sort = "F_LOJA" Else ADOREG.Recordset.Sort = "F_NOME"
        
        ADOREG.Recordset.MoveFirst
        ADOREG.Recordset.Find "m_f_cod = " & w_cod_atual, , adSearchForward
        
        If ADOREG.Recordset.EOF Then
            Op_Click 5
            ADOREG.Recordset.Find "m_f_cod = " & w_cod_atual, , adSearchForward
            If ADOREG.Recordset.EOF Then
                MsgBox "Não existe ficha desse funcionário no Mês " & txt_PMes & " em nenhuma loja!", vbCritical, "Ficha não encontrada"
                ADOREG.Recordset.MoveFirst
                w_Op = w_Op = -Op(1).value * 1 - Op(2).value * 2 - Op(3).value * 3 - Op(4).value * 4 - Op(5).value * 5 - Op(6).value * 6 - Op(7).value * 7
                Op_Click (5)
                Op_Click (w_Op)
                cmdFiltrar_Click
            End If
            'Set ADOREG.Recordset = de.rscmdSqlVisAltFichas.Clone
            'If optLoja.value Then ADOREG.Recordset.Sort = "F_LOJA" Else ADOREG.Recordset.Sort = "F_NOME"
        End If
    
    End If
    
    w_reset_tipo = True
    
End Sub

Sub bt_VoltarDT_Click()
On Error Resume Next
    Dim w_cod_atual As String
    
    w_cod_atual = txt_F_COD
    
    w_reset_tipo = False
    
    If CDbl(txt_PMes) = 1 Then
        txt_PMes = 12
        txt_PAno = CDbl(txt_PAno) - 1
    Else
        txt_PMes = Format(CDbl(txt_PMes) - 1, "00")
    End If
    
    wID = W_INDEX
    W_INDEX = 2
    Call cmdFiltrar_Click
    
    W_INDEX = wID
    Call cmdFiltrar_Click
    
    Op(2).Caption = Format(CDbl(txt_PMes), "00") & "/" & txt_PAno
    
    If optLoja.value Then ADOREG.Recordset.Sort = "F_LOJA" Else ADOREG.Recordset.Sort = "F_NOME"
    
    ADOREG.Recordset.MoveFirst
    ADOREG.Recordset.Find "m_f_cod = " & w_cod_atual, , adSearchForward
    
    If ADOREG.Recordset.EOF Then
        Op_Click 5
        ADOREG.Recordset.Find "m_f_cod = " & w_cod_atual, , adSearchForward
        If ADOREG.Recordset.EOF Then
            MsgBox "Não existe ficha desse funcionário no Mês " & txt_PMes & " em nenhuma loja!", vbCritical, "Ficha não encontrada"
            ADOREG.Recordset.MoveFirst
            w_Op = -Op(1).value * 1 - Op(2).value * 2 - Op(3).value * 3 - Op(4).value * 4 - Op(5).value * 5 - Op(6).value * 6 - Op(7).value * 7
            Op_Click (5)
            Op_Click (w_Op)
            cmdFiltrar_Click
        End If
        'Op_Click 1
        'txt_Pesq = TXT_LOGO2
        'cmdFiltrar_Click
        'adoReg.Recordset.Find "m_f_cod = " & w_cod_atual, , adSearchForward
        'MsgBox "Este funcionário não possui ficha nesta mesma loja no mês " & txt_PMes & ", mas possui na loja " & TXT_LOGO & "!", vbCritical, "Ficha encontrada em outra loja"
        'Set adoReg.Recordset = de.rscmdSqlVisAltFichas.Clone
        'If optLoja.value Then adoReg.Recordset.Sort = "F_LOJA" Else adoReg.Recordset.Sort = "F_NOME"
    End If
    
    w_reset_tipo = True
    
End Sub

Private Sub btnAcessoEspecial_Click()
    frm_Acesso_Especial.Show 1
End Sub

Private Sub btRptADM_Click()
On Error GoTo err1

    If de.rscmdSqlFichaMensalADM_Grouping.State = 1 Then de.rscmdSqlFichaMensalADM_Grouping.Close
    w_DtIni = CVDate("01/" & Format(TXT_MES, "00") & "/" & TXT_ANO)
    If CDbl(TXT_MES) + 1 > 12 Then
        w_DtFim = CVDate("01/01/" & TXT_ANO + 1) - 1
    Else
        w_DtFim = CVDate("01/" & Format(CDbl(TXT_MES) + 1, "00") & "/" & TXT_ANO) - 1
    End If
    
    rptFichaMensalADM.Sections("SecCab").Controls("lbAno").Caption = TXT_ANO
    rptFichaMensalADM.Sections("SecCab").Controls("lbMes").Caption = TXT_MES
    
    de.cmdSqlFichaMensalADM_Grouping w_DtIni, w_DtFim, TXT_MES, TXT_ANO
    
    rptFichaMensalADM.Show
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub btRptDem_Click()
On Error GoTo err1

    If de.rscmdSqlFichaMensalDem_Grouping.State = 1 Then de.rscmdSqlFichaMensalDem_Grouping.Close
    de.cmdSqlFichaMensalDem_Grouping TXT_MES, TXT_ANO
    
    
    rptFichaMensalDem.Sections("SecCab").Controls("lbTitulo").Caption = "(D)  " & Format(TXT_MES, "00") & " / " & TXT_ANO
    'rptFichaMensalDem.Sections("SecCab").Controls("lbData").Caption = Format(Date, "DD/MM/YY") & " " & Format(Time, "hh:mm") & "hs"
    
    rptFichaMensalDem.Show
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub btRptREG_Click()
On Error GoTo err1

    If de.rscmdSqlFichaMensalREG_Grouping.State = 1 Then de.rscmdSqlFichaMensalREG_Grouping.Close
    w_DtIni = CVDate("01/" & Format(TXT_MES, "00") & "/" & TXT_ANO)
    If CDbl(TXT_MES) + 1 > 12 Then
        w_DtFim = CVDate("01/01/" & TXT_ANO + 1) - 1
    Else
        w_DtFim = CVDate("01/" & Format(CDbl(TXT_MES) + 1, "00") & "/" & TXT_ANO) - 1
    End If
    
    de.cmdSqlFichaMensalREG_Grouping TXT_MES, TXT_ANO, w_DtIni, w_DtFim
    
    rptFichaMensalREG.Show
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub btRptVctoFerias_Click()
On Error GoTo err1
    
    Op(7).value = True

   ' If W_FILTRO <> "" Then
        If de.rscmdSqlVctoFerias.State = 1 Then de.rscmdSqlVctoFerias.Close
        de.cmdSqlVctoFerias TXT_ANO, TXT_MES
       ' de.rscmdSqlVctoFerias.Filter = IIf(W_FILTRO = "", "m_mes = 0", W_FILTRO)
            
        Dim w_Valor As Variant
        
           If de.rscmdSqlVctoFerias.RecordCount > 0 Then
           
                de.rscmdSqlVctoFerias.MoveFirst
                Do While Not de.rscmdSqlVctoFerias.EOF
                    If de.cnc.Execute("Select c_valor FROM TAB_DESC_CALC WHERE c_n_ficha = " & de.rscmdSqlVctoFerias.Fields("m_nficha") & " AND c_tp_conta = 24").RecordCount > 0 Then
                        w_Valor = CDbl(de.cnc.Execute("Select SUM(c_valor) FROM TAB_DESC_CALC WHERE c_n_ficha = " & de.rscmdSqlVctoFerias.Fields("m_nficha") & " AND c_tp_conta = 24").Fields(0))
                    Else
                        w_Valor = 0
                    End If
                    de.rscmdSqlVctoFerias.Fields("ValorPG") = CDbl(w_Valor)
                    de.rscmdSqlVctoFerias.MoveNext
                Loop
                
            End If
       
        de.rscmdSqlVctoFerias.Sort = "ValorPG"
        rptVctoFerias.Sections("SecCab").Controls("lbMes").Caption = TXT_MES & " / " & TXT_ANO
        rptVctoFerias.Show
    'End If


sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub chkGerentes_Click()
    If chkGerentes.value And chkOutros.value Then
        Op_Click 5
    ElseIf chkGerentes.value Then
        txt_Pesq = "'G'"
        FILTRAR 8
    ElseIf chkOutros.value Then
        txt_Pesq = "'G'"
        FILTRAR 9
    ElseIf chkGerentes.value = 0 And chkOutros.value = 0 Then
        Op_Click 5
    End If
    
    txt_Pesq = ""
    
End Sub

Private Sub chkOutros_Click()
    If chkGerentes.value And chkOutros.value Then
        Op_Click 5
    ElseIf chkGerentes.value Then
        txt_Pesq = "'G'"
        FILTRAR 8
    ElseIf chkOutros.value Then
        txt_Pesq = "'G'"
        FILTRAR 9
    ElseIf chkGerentes.value = 0 And chkOutros.value = 0 Then
        Op_Click 5
    End If
    
    txt_Pesq = ""
    
End Sub


Private Sub cbMostrar_Click()
        cmdMostrar_Click
End Sub

Private Sub cbMostrar_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then
        cmdMostrar_Click
    End If
End Sub

Sub CK_ACF_Click()
If Not ADOREG.Recordset.EOF Then
    If (CK_ACF.Enabled = True And TXT_NFICHA = ADOREG.Recordset.Fields("M_NFICHA")) Or w_bloq Then
        If BarraF.Buttons("salvar").Enabled = False Then Editar 0

        If CK_ACF.value = 0 Then
            'TXT_AC_F = ""
            'ADOREG.Recordset.Fields("m_dt_acf") = Null
        Else
            If CK_DEM.value = 0 Then CK_DEM.value = 1
            'If TXT_AC_F.Text = "" Or TXT_AC_F.Text = Null Then
                w_ac_f = Format(Date, "DD/MM/YYYY") & " " & Format(Time, "hh:mm") & ": " & TXT_AC_F
                TXT_AC_F = w_ac_f
            'End If
            On Error Resume Next
            TXT_AC_F.SetFocus
            Sendkeys "{end}"
        End If
        Salvar
        If BarraF.Buttons("salvar").Enabled = True Then Salvar
    End If
End If

End Sub




Private Sub CK_DEM_Click()
'    If W_CK_DEM = True And adoReg.Recordset.Fields("M_BLOQ") = False And CK_DEM.Enabled = True And UCase(frmLogin.txtUserName) = UCase(NomeMestre) Then

    If (UCase(frmLogin.txtUserName) = UCase(NomeMestre) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre2) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre3)) Then
        
     If Not ADOREG.Recordset.EOF Then
        If IsNull(ADOREG.Recordset.Fields("M_DT_DEM")) And CK_DEM.value = 1 Then
            'MsgBox "VOCÊ NÃO PODE VISTAR ,  SEM DATA DE DEMISSÃO!", vbCritical
            CK_DEM.value = 0
        ElseIf ADOREG.Recordset.Fields("M_F_COD") = txt_F_COD Then
            If BarraF.Buttons("salvar").Enabled = False Then Editar 0
            '*** OK DT_DEM
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_DEM_OK = " & CK_DEM.value * -1 & "  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        
            '*** SÓ EDITA SE AINDA NÃO FOI CHECADO   ***
            If CK_DEM = 1 Then
                CK_ACF.Enabled = True
                'TXT_AC_F.Enabled = True
            End If
            Salvar
        End If
      End If
    End If
End Sub


Private Sub cmd13_Click()
On Error Resume Next
    Dim w_Dt13 As Date
    Dim w_DtDif As Integer
    Dim w_Vr13, w_Piso13 As Double
    Dim w_Desc13, w_obs13, w_dt13fim As String
    
    ADOREG.Recordset.MoveFirst
    Do While Not ADOREG.Recordset.EOF
        If IsNull(ADOREG.Recordset.Fields("m_dt_acf")) Then
            w_Dt13 = "01/01/2000"
            w_DtDif = "0"
            w_Vr13 = 0
            w_Piso13 = 0
            w_Desc13 = ""
            w_obs13 = ""
            w_dt13fim = ""
            
            If ADOREG.Recordset.Fields("m_dt_reg") = "" Or IsNull(ADOREG.Recordset.Fields("m_dt_reg")) Then
                w_Dt13 = CVDate(ADOREG.Recordset.Fields("m_dt_adm"))
            Else
                w_Dt13 = CVDate(ADOREG.Recordset.Fields("m_dt_reg"))
            End If
                
            w_DtDif = DateDiff("m", w_Dt13, CVDate("31/12/" & Year(Date)))
                
            If (DateDiff("d", Day(w_Dt13), Day((UltDiaMes(Month(w_Dt13), Year(w_Dt13))))) + 1) >= 15 Then
                w_DtDif = w_DtDif + 1
            End If
            
            If w_DtDif >= 12 Then w_DtDif = 12
            
            w_Piso13 = 0
            w_Piso13 = de.cnc.Execute("SELECT F_VPISO_R FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
            If w_Piso13 = 0 Then
                w_Piso13 = de.cnc.Execute("SELECT F_VPISO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
            End If
                  
            If w_Piso13 > 0 Then
                w_Vr13 = (w_Piso13 / 12) * w_DtDif
            Else
                w_Vr13 = 0
            End If
            
            
            
            If ADOREG.Recordset.Fields("M_TIPO") = "G" Then
                w_Desc13 = w_DtDif & "/12 13º"
            Else
                w_Desc13 = w_DtDif & "/12 13º   |   (" & Format(w_Piso13, "####0.00") & " / 12 = " & Format(w_Piso13 / 12, "####0.00") & ") * " & w_DtDif & " = " & Format(w_Vr13, "####0.00")
            End If
            
            
            w_obs13 = "13º/" & Year(Date) & " OK   |   " & w_DtDif & "/12"
            w_dt13fim = CVDate("31/12/" & Year(Date))
            
            de.cnc.Execute ("DELETE FROM TAB_DESC_CALC Where (C_TP_CONTA = 32) And (C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFicha") & ")")
            de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("m_nficha"), 32, "+", w_Vr13, w_Desc13, ADOREG.Recordset.Fields("m_logo"), 0, 0, 0, ADOREG.Recordset.Fields("m_f_cod")
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_13_OK = 0 , F_13_ULT_PG = F_13_PG, F_13_PG = '" & w_dt13fim & "' , F_13_OBS = '" & w_obs13 & "' WHERE (F_Codigo = " & ADOREG.Recordset.Fields("M_F_COD") & ")"
            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_13_OK = 0 , M_13_ULT_PG = M_13_PG, M_13_PG = '" & w_dt13fim & "', M_13_OBS = '" & w_obs13 & "'  WHERE (M_F_Cod = " & ADOREG.Recordset.Fields("M_F_COD") & ")"
        End If
        ADOREG.Recordset.MoveNext
    Loop
    
    ADOREG.Recordset.MoveFirst
        
        
    'Dados Contas
    Lancamentos
    
    
sair:
    Exit Sub
err1:
    
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    
Resume sair
    
End Sub

Private Sub cmdAddLanç_SalF_Click()
On Error GoTo err1
    Dim wSalFam
    
    wSalFam = de.cnc.Execute("Select Sal_Familia from tab_config").Fields(0)
    
    wValor = 0
    wValor = Format(ADOREG.Recordset.Fields("m_num_filhos") * wSalFam, "0.00")  'Calcula Salario
    wDesc = "(" & Format(wSalFam, "0.00") & " x " & ADOREG.Recordset.Fields("m_num_filhos") & ") = " & Format(wValor, "0.00")
    de.cnc.Execute ("DELETE FROM TAB_DESC_CALC Where (C_TP_CONTA = 26) And (C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFicha") & ")")
    de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFicha"), 26, "+", wValor, wDesc, "", "0", "0", "0", "0"
    
    'Dados Contas
    Lancamentos
    
sair:
    Exit Sub
err1:
    
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    
    Resume sair
End Sub

Private Sub cmdAddLanç_SalFTodos_Click()
On Error GoTo err1
    Dim wSalFam
    
    ADOREG.Recordset.MoveFirst
    Do While Not ADOREG.Recordset.EOF
        
        wSalFam = de.cnc.Execute("Select Sal_Familia from tab_config").Fields(0)
        
        wValor = 0
        
        If ADOREG.Recordset.Fields("m_num_filhos") > 0 Then
            wValor = Format(ADOREG.Recordset.Fields("m_num_filhos") * wSalFam, "0.00")  'Calcula Salario
            wDesc = "(" & Format(wSalFam, "0.00") & " x " & ADOREG.Recordset.Fields("m_num_filhos") & ") = " & Format(wValor, "0.00")
            de.cnc.Execute ("DELETE FROM TAB_DESC_CALC Where (C_TP_CONTA = 26) And (C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFicha") & ")")
            de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFicha"), 26, "+", wValor, wDesc, "", "0", "0", "0", "0"
        End If
        
        ADOREG.Recordset.MoveNext
    Loop
    
    ADOREG.Recordset.MoveFirst
        
    'Dados Contas
    Lancamentos
    
sair:
    Exit Sub
err1:
    
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    
    Resume sair
End Sub

Private Sub cmdAddSalario_Click()

    If Not isMesValido(txt_F_COD, TXT_MES, TXT_ANO) Then 'Verifica se é mês atual ou passado
        If MsgBox("Você está alterando uma ficha que NÃO É DO MÊS ATUAL. Deseja continuar mesmo assim?", vbYesNo, "Alteração de ficha") = vbNo Then
            Exit Sub
        End If
        If ADOREG.Recordset.Fields("M_BLOQ") Then
            MsgBox "Esta ficha não é do mês atual e está BLOQUEADA!", vbCritical
            Exit Sub
        End If
    End If
    
    If MsgBox("Deseja (re)gerar a comissão do(a) funcionário(a) " & TXT_FUNC & "?", vbYesNo, "Gerar comissão") = vbNo Then
        Exit Sub
    End If
    
    
    Select Case ADOREG.Recordset.Fields("M_TIPO")
        
    Case "V": 'VENDEDOR
         Dim dtIni, dtFim As Date
        Dim adoComis As ADODB.Recordset
        Dim adoComisTMP As ADODB.Recordset
        Dim adoComisDump As ADODB.Recordset
        Dim sqlComissTMP As String
        Dim vrMeta, vrMetaBonus, descMeta, vrSalario, wFixoMeta
        Dim w_ado_venda As ADODB.Recordset
        
            frm_ESCOLHA_DATA.Show 1
            
            dtIni = CDate(frm_ESCOLHA_DATA.TXT_DT_INICIAL)
            dtFim = CDate(frm_ESCOLHA_DATA.TXT_DT_FINAL)
           
            If de.rscmdComiss_Grouping.State = 1 Then de.rscmdComiss_Grouping.Close
            
            'On Error Resume Next
            'de.cmdDROPtmpComis1
            'de.cmdDROPtmpComis2
            
            'de.cmdCREATEtmpComis1
            'de.cmdCREATEtmpComis2
            
            de.cmdDELETEtmpComis1
            de.cmdDELETEtmpComis2
            
            de.cmdAddtmpComis1 dtIni, dtFim, dtIni, dtFim, dtIni, dtFim, dtIni, dtFim, dtIni, dtFim
            de.cmdAddtmpComis2 dtIni, dtFim, dtIni, dtFim
                
            de.cmdComiss_Grouping
            
            If MsgBox("Tem certeza que deseja (RE)GERAR A COMISSÃO DE VENDEDORES para " & lblMes.Caption & "?", vbYesNo, "GERAR COMISSÃO") = vbYes Then
            
                If de.rscmdComissGerar.State = 1 Then de.rscmdComissGerar.Close
                de.cmdComissGerar
                Set adoComis = de.rscmdComissGerar.Clone
                    
                Dim w_Dt, w_dtUlt As Date
                Dim w_DtDiff, w_ultDiaMes As Integer
                Dim w_Piso, w_Comis, w_Premio, w_PisoOriginal, w_FixoOriginal, w_Fixo As Double
                    
                'Do While Not adoReg.Recordset.EOF
                
                   'If adoReg.Recordset.Fields("M_TIPO") = "V" Then
                
                    'If adoReg.Recordset.Fields("M_TIPO") = "V" And ( _
                    '        adoReg.Recordset.Fields("M_LOGO") = "FH") Then
                    '        adoReg.Recordset.Fields("M_LOGO") = "BA" Or _
                    '        adoReg.Recordset.Fields("M_LOGO") = "AC" Or _
                    '        adoReg.Recordset.Fields("M_LOGO") = "SV" Or _
                    '        adoReg.Recordset.Fields("M_LOGO") = "LB" Or _
                    '        adoReg.Recordset.Fields("M_LOGO") = "L4" Or _
                    '        adoReg.Recordset.Fields("M_LOGO") = "SI" Or _
                    '        adoReg.Recordset.Fields("M_LOGO") = "SY" Or _
                    '        adoReg.Recordset.Fields("M_LOGO") = "CJ") Then
        
                
                    '**** BONUS DE META
                    'COD da loja do cx do ANO atual
                    'Set w_ado_venda = de.cnc.Execute("SELECT TAB_VENDA.V_VR From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND TAB_FUNCIONARIO.F_DT_DEM IS NULL AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & Format(adoReg.Recordset.Fields("M_MES"), "00") & " AND Right(TAB_VENDA.V_DATA,4) = " & adoReg.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & adoReg.Recordset.Fields("M_F_COD")).Clone
                    Set w_ado_venda = de.cnc.Execute("SELECT TAB_VENDA.V_VR FROM TAB_VENDA, TAB_FUNCIONARIO INNER JOIN Lojb010 ON TAB_FUNCIONARIO.F_Cod_L = Lojb010.COD_LOJ WHERE (((TAB_FUNCIONARIO.F_DT_DEM) Is Null) AND ((Right(Left([TAB_VENDA].[V_DATA],5),2))=" & Format(ADOREG.Recordset.Fields("M_MES"), "00") & ") AND ((Right([TAB_VENDA].[V_DATA],4))=" & ADOREG.Recordset.Fields("M_ANO") & ") AND ((TAB_FUNCIONARIO.F_Codigo)=" & ADOREG.Recordset.Fields("M_F_COD") & ") AND ((Lojb010.NUM)=Right([tab_venda].[v_f_loja],2)));").Clone
                    If Not w_ado_venda.EOF Then
                        vrVenda = w_ado_venda.Fields(0)
                    Else
                        vrVenda = 0
                        'MsgBox "Não há lançamentos do logo " & adoReg.Recordset.Fields("M_LOGO") & " para o período: " & Format(adoReg.Recordset.Fields("M_MES"), "00") & " / " & adoReg.Recordset.Fields("M_ANO") & "! Ignorando...", vbCritical
                    End If
            
                    vrMeta = 999
                    If ADOREG.Recordset.Fields("M_LOGO") <> "XX" And ADOREG.Recordset.Fields("M_LOGO") <> "99" Then
                        vrMeta = de.cnc.Execute("SELECT TAB_META.MT_VALOR FROM TAB_META, Lojb010 INNER JOIN TAB_FUNCIONARIO ON Lojb010.COD_LOJ = TAB_FUNCIONARIO.F_Cod_L WHERE (((TAB_META.MT_MES)=" & ADOREG.Recordset.Fields("M_MES") & ") AND ((TAB_META.MT_ANO)=" & ADOREG.Recordset.Fields("M_ANO") & ") AND ((TAB_FUNCIONARIO.F_Codigo)=" & ADOREG.Recordset.Fields("M_F_COD") & ") AND ((Lojb010.NUM)=Right([tab_meta].[mt_f_loja],2)))").Fields(0)
                    End If
                    
                    '****
                    
                    If ADOREG.Recordset.Fields("F_COD_CENTRAL") <> "" And IsNull(ADOREG.Recordset.Fields("M_DT_DEM")) Then
        
                        
                        adoComis.Filter = "F_4023717930 = " & ADOREG.Recordset.Fields("F_COD_CENTRAL")
                        If Not adoComis.EOF Then
                            w_Piso = 0
                            w_Piso = de.cnc.Execute("SELECT F_VPISO_R FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
                            If w_Piso = 0 Then
                                w_Piso = de.cnc.Execute("SELECT F_VPISO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
                            End If
                            
                            w_Comis = 0
                            w_Premio = 0
                            w_Comis = CDbl(adoComis.Fields("COMTOTAL"))
                            w_Premio = CDbl(adoComis.Fields("F_1373503546"))
                            
                            de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & " AND (C_TP_CONTA = 20 OR C_TP_CONTA = 21 OR C_TP_CONTA = 23 OR C_TP_CONTA = 102 OR C_TP_CONTA = 101)")
                            
                        'If adoReg.Recordset.Fields("M_MES") > "12" Then
                        '        w_Fixo = de.cnc.Execute("SELECT IIf(IsNull(Lojb010.FixoVend),0,Lojb010.FixoVend) AS Fixo, TAB_FICHA_MENS.M_MES, TAB_FICHA_MENS.M_ANO FROM TAB_FICHA_MENS INNER JOIN Lojb010 ON TAB_FICHA_MENS.M_LOGO = Lojb010.COD_LOJ " & _
                        '             " WHERE (((TAB_FICHA_MENS.M_F_COD)=" & adoReg.Recordset.Fields("M_F_COD") & ") " & _
                        '             " AND ((TAB_FICHA_MENS.M_MES)=" & adoReg.Recordset.Fields("M_MES") & ") AND ((TAB_FICHA_MENS.M_ANO)=" & adoReg.Recordset.Fields("M_ANO") & "));").Fields(0)
                        'Else
                            w_Fixo = 0
                        'End If
                            
                          If (w_Comis + w_Premio + w_Fixo) <= w_Piso Then
                             'w_ultDiaMes = 30
                             w_ultDiaMes = Day(UltDiaMes(ADOREG.Recordset.Fields("m_mes"), ADOREG.Recordset.Fields("m_ano")))
                             If w_ultDiaMes < 30 Then w_ultDiaMes = 30
                         
                             'If adoReg.Recordset.Fields("m_dt_reg") = "" Or IsNull(adoReg.Recordset.Fields("m_dt_reg")) Then
                                 w_Dt = CVDate(ADOREG.Recordset.Fields("m_dt_adm"))
                             'Else
                             '    w_Dt = CVDate(adoReg.Recordset.Fields("m_dt_reg"))
                             'End If
                             
                             
                             If IsDate(ADOREG.Recordset.Fields("M_DT_DEM")) Then
                                 w_dtUlt = CVDate(ADOREG.Recordset.Fields("M_DT_DEM"))
                             Else
                                 w_dtUlt = CVDate(UltDiaMes(ADOREG.Recordset.Fields("m_mes"), ADOREG.Recordset.Fields("m_ano")))
                                 'If Day(w_dtUlt) = 31 Then w_dtUlt = w_dtUlt - 1
                                 If Day(w_dtUlt) = 28 Then w_dtUlt = w_dtUlt + 2
                                 If Day(w_dtUlt) = 29 Then w_dtUlt = w_dtUlt + 1
                             End If
                             
                             'If Month(w_Dt) < Month(w_dtUlt) Then w_Dt = CVDate("01/" & Month(w_dtUlt) & "/" & Year(w_dtUlt))
                             
                             w_DtDiff = DateDiff("d", w_Dt, w_dtUlt) + 1
                             
                             w_PisoOriginal = w_Piso
                             w_FixoOriginal = w_Fixo
                             'MsgBox "Diff: " & w_DtDiff & " - Ini: " & w_Dt & " - Final: " & w_dtUlt
                             If (w_DtDiff < w_ultDiaMes) And (w_DtDiff < 30) Then
                                 If w_ultDiaMes < 30 Then
                                    If w_Fixo > 0 Then w_Fixo = ((w_Fixo / w_ultDiaMes) * w_DtDiff)
                                    w_Piso = ((w_Piso / w_ultDiaMes) * w_DtDiff)
                                 Else
                                    If w_Fixo > 0 Then w_Fixo = ((w_Fixo / 30) * w_DtDiff)
                                    w_Piso = ((w_Piso / 30) * w_DtDiff)
                                    w_ultDiaMes = 30
                                 End If
                                 
                                 If w_FixoOriginal > 0 Then
                                    If (w_Comis + w_Premio + w_Fixo) <= w_Piso Then
                                        de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 23, "+", Format(w_Piso, "0.00"), "#NÃO ATINGIU O PISO# Fixo: " & Format(w_FixoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_FixoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Fixo, "0.00") & " + Comissão: " & Format(w_Comis, "0.00") & " + Prêmio: " & Format(w_Premio, "0.00") & " = " & Format(w_Fixo + w_Comis + w_Premio, "0.00") & ". Piso: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                                    Else
                                        de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 20, "+", Format(w_Comis, "0.00"), "#SALÁRIO MAIOR QUE PISO PROPORCIONAL# Comissão: " & Format(w_Comis, "0.00") & ". Piso: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                                        de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 102, "+", Format(w_Fixo, "0.00"), "#SALÁRIO MAIOR QUE PISO PROPORCIONAL# Fixo: " & Format(w_FixoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_FixoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Fixo, "0.00") & ". Piso: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                                        de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 21, "+", Format(w_Premio, "0.00"), "PRÊMIO [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                                    End If
                                 Else
                                    If (w_Comis + w_Premio) <= w_Piso Then
                                        de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 23, "+", Format(w_Piso, "0.00"), "#NÃO ATINGIU O PISO# Comissão: " & Format(w_Comis, "0.00") & " + Prêmio: " & Format(w_Premio, "0.00") & " = " & Format(w_Comis + w_Premio, "0.00") & ". Piso: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                                    Else
                                        de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 20, "+", Format(w_Comis, "0.00"), "#COMISSÃO MAIOR QUE PISO PROPORCIONAL# Comissão: " & Format(w_Comis, "0.00") & " + Prêmio: " & Format(w_Premio, "0.00") & " = " & Format(w_Comis + w_Premio, "0.00") & ". Piso: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                                        de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 21, "+", Format(w_Premio, "0.00"), "PRÊMIO [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                                    End If
                                 End If
                
                             Else
                                If w_FixoOriginal > 0 Then
                                    de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 23, "+", Format(w_Piso, "0.00"), "#NÃO ATINGIU O PISO# Fixo: " & Format(w_Fixo, "0.00") & " + Comissão: " & Format(w_Comis, "0.00") & " + Prêmio: " & Format(w_Premio, "0.00") & " = " & Format(w_Fixo + w_Comis + w_Premio, "0.00") & ". Piso: " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                                Else
                                    de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 23, "+", Format(w_Piso, "0.00"), "#NÃO ATINGIU O PISO# Comissão: " & Format(w_Comis, "0.00") & " + Prêmio: " & Format(w_Premio, "0.00") & " = " & Format(w_Comis + w_Premio, "0.00") & ". Piso: " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                                End If
                             End If
                             
                         Else
                            If w_Fixo > 0 Then
                               de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 20, "+", Format(w_Comis, "0.00"), "COMISSÃO [GERADO AUTOMATICAMENTE] Comissão: " & Format(w_Comis, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                               de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 102, "+", Format(w_Fixo, "0.00"), "FIXO [GERADO AUTOMATICAMENTE] Fixo: " & Format(w_Fixo, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                               de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 21, "+", Format(w_Premio, "0.00"), "PRÊMIO [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                            Else
                               de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 20, "+", Format(w_Comis, "0.00"), "COMISSÃO [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                               de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 21, "+", Format(w_Premio, "0.00"), "PRÊMIO [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                            End If
                         End If
                    End If
                   ElseIf IsNull(ADOREG.Recordset.Fields("M_DT_DEM")) Then
                   
                    'Dim w_Dt, w_dtUlt As Date
                    'Dim w_DtDiff, w_ultDiaMes As Integer
                    'Dim w_Piso, w_Comis, w_Premio, w_PisoOriginal As Double
                        
                        If Not adoComis.EOF Then
                            w_Piso = 0
                            w_Piso = de.cnc.Execute("SELECT F_VPISO_R FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
                            If w_Piso = 0 Then
                                w_Piso = de.cnc.Execute("SELECT F_VPISO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
                            End If
                            
                            'If adoReg.Recordset.Fields("M_MES") = "1" Or adoReg.Recordset.Fields("M_MES") = "2" Or adoReg.Recordset.Fields("M_MES") = "3" Then
                            '    w_Fixo = de.cnc.Execute("SELECT IIf(IsNull(Lojb010.FixoVend),0,Lojb010.FixoVend) AS Fixo, TAB_FICHA_MENS.M_MES, TAB_FICHA_MENS.M_ANO FROM TAB_FICHA_MENS INNER JOIN Lojb010 ON TAB_FICHA_MENS.M_LOGO = Lojb010.COD_LOJ " & _
                            '         " WHERE (((TAB_FICHA_MENS.M_F_COD)=" & adoReg.Recordset.Fields("M_F_COD") & ") " & _
                            '         " AND ((TAB_FICHA_MENS.M_MES)=" & adoReg.Recordset.Fields("M_MES") & ") AND ((TAB_FICHA_MENS.M_ANO)=" & adoReg.Recordset.Fields("M_ANO") & "));").Fields(0)
                            'Else
                                w_Fixo = 0
                            'End If
                            
                            de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & " AND (C_TP_CONTA = 20 OR C_TP_CONTA = 21 OR C_TP_CONTA = 23  OR C_TP_CONTA = 102)")
                            
                             'w_ultDiaMes = 30
                             w_ultDiaMes = Day(UltDiaMes(ADOREG.Recordset.Fields("m_mes"), ADOREG.Recordset.Fields("m_ano")))
                             If w_ultDiaMes < 30 Then w_ultDiaMes = 30
                         
                             'If adoReg.Recordset.Fields("m_dt_reg") = "" Or IsNull(adoReg.Recordset.Fields("m_dt_reg")) Then
                                 w_Dt = CVDate(ADOREG.Recordset.Fields("m_dt_adm"))
                            ' Else
                             '    w_Dt = CVDate(adoReg.Recordset.Fields("m_dt_reg"))
                            ' End If
                             
                             
                             If IsDate(ADOREG.Recordset.Fields("M_DT_DEM")) Then
                                 w_dtUlt = CVDate(ADOREG.Recordset.Fields("M_DT_DEM"))
                             Else
                                 w_dtUlt = CVDate(UltDiaMes(ADOREG.Recordset.Fields("m_mes"), ADOREG.Recordset.Fields("m_ano")))
                                 'If Day(w_dtUlt) = 31 Then w_dtUlt = w_dtUlt - 1
                                 If Day(w_dtUlt) = 28 Then w_dtUlt = w_dtUlt + 2
                                 If Day(w_dtUlt) = 29 Then w_dtUlt = w_dtUlt + 1
                             End If
                             
                             'If Month(w_Dt) < Month(w_dtUlt) Then w_Dt = CVDate("01/" & Month(w_dtUlt) & "/" & Year(w_dtUlt))
                             
                             w_DtDiff = DateDiff("d", w_Dt, w_dtUlt) + 1
                             
                             w_PisoOriginal = w_Piso
                             w_FixoOriginal = w_Fixo
                             If (w_DtDiff < w_ultDiaMes) And (w_DtDiff < 30) Then
                                 If w_ultDiaMes < 30 Then
                                    If w_Fixo > 0 Then w_Fixo = ((w_Fixo / w_ultDiaMes) * w_DtDiff)
                                    w_Piso = ((w_Piso / w_ultDiaMes) * w_DtDiff)
                                 Else
                                    If w_Fixo > 0 Then w_Fixo = ((w_Fixo / 30) * w_DtDiff)
                                    w_Piso = ((w_Piso / 30) * w_DtDiff)
                                    w_ultDiaMes = 30
                                 End If
                                If w_FixoOriginal > 0 Then
                                   de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 23, "+", Format(w_Piso, "0.00"), "#NÃO ATINGIU O PISO PROPORCIONAL# Fixo: " & Format(w_Fixo, "0.00") & " + Comissão: " & Format(0, "0.00") & " + Prêmio: " & Format(0, "0.00") & " = " & Format(0 + 0, "0.00") & ". Piso: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                                Else
                                   de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 23, "+", Format(w_Piso, "0.00"), "#NÃO ATINGIU O PISO PROPORCIONAL# Comissão: " & Format(0, "0.00") & " + Prêmio: " & Format(0, "0.00") & " = " & Format(0 + 0, "0.00") & ". Piso: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                                End If
                             
                             Else
                                If w_FixoOriginal > 0 Then
                                    de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 23, "+", Format(w_Piso, "0.00"), "#NÃO ATINGIU O PISO# Fixo: " & Format(w_Fixo, "0.00") & " + Comissão: " & Format(0, "0.00") & " + Prêmio: " & Format(0, "0.00") & " = " & Format(0 + 0, "0.00") & ". Piso: " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                                Else
                                    de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 23, "+", Format(w_Piso, "0.00"), "#NÃO ATINGIU O PISO# Comissão: " & Format(0, "0.00") & " + Prêmio: " & Format(0, "0.00") & " = " & Format(0 + 0, "0.00") & ". Piso: " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                                End If
                             End If
                    End If
                    
                   End If
                   
                     vrSalario = w_Comis + w_Premio + w_Fixo
                     vrMetaBonus = 0
                     wFixoMeta = 0
                     wPercMeta = 0
                        
                    If IsNull(ADOREG.Recordset.Fields("M_DT_DEM")) Then
                     
                        If vrVenda >= (vrMeta) Then
                        
                            If ADOREG.Recordset.Fields("m_mes") = 1 Then
        
                                If w_Piso > vrSalario Then
                                   vrSalario = w_Piso
                                End If
                            End If
                        
                            wFixoMeta = 150
                            If (((vrVenda / vrMeta) - 1) * 100) >= 5 Then
                                vrMetaBonus = vrSalario * 0.05
                                wPercMeta = 5
                            Else
                                vrMetaBonus = vrSalario * ((vrVenda / vrMeta) - 1)
                                wPercMeta = ((vrVenda / vrMeta) - 1) * 100
                            End If
        
                            de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & " AND (C_TP_CONTA = 101)")
                            'Bonus de % do Salário
                            de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 101, "+", Format(vrMetaBonus + wFixoMeta, "0.00"), "#BONUS DE META# Fixo(" & Format(w_Fixo, "0.00") & ") + Comissão(" & Format(w_Comis, "0.00") & ") + Prêmio(" & Format(w_Premio, "0.00") & ") = " & Format(vrSalario, "0.00") & " * " & Format(wPercMeta, "0.0") & "% = " & Format(vrMetaBonus, "0.00") & " + Fixo Meta(" & Format(wFixoMeta, "0.00") & ") = " & Format(vrMetaBonus + wFixoMeta, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                            vrSalario = vrSalario + vrMetaBonus + wFixoMeta
                        Else
                           descMeta = ""
                        End If
                        
        
                     
                    End If
                   
                   End If
                   
                   'End If
        
                   'adoReg.Recordset.MoveNext
                   
                    
               ' Loop
                
                'adoReg.Recordset.MoveFirst
                Lancamentos
                

                    
        
        
    Case "G": 'GERENTE
        Dim vrFixo, vrMinimo, percComis, vrComis
        Dim vrVendaAnt, percVenda, vrBonus, vrSalarioBonus
            
 
             'COD da loja do gerente do ANO anterior
            Set w_ado_vendaAnt = de.cnc.Execute("SELECT TAB_VENDA.V_VR From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND TAB_FUNCIONARIO.F_DT_DEM IS NULL AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & Format(ADOREG.Recordset.Fields("M_MES"), "00") & " AND Right(TAB_VENDA.V_DATA,4) = " & CInt(ADOREG.Recordset.Fields("M_ANO")) - 1 & " AND TAB_FUNCIONARIO.F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD")).Clone
            If Not w_ado_vendaAnt.EOF Then
                vrVendaAnt = w_ado_vendaAnt.Fields(0)
            Else
                MsgBox "Não há lançamentos do logo " & ADOREG.Recordset.Fields("M_LOGO") & " para o período: " & Format(ADOREG.Recordset.Fields("M_MES"), "00") & " / " & CInt(ADOREG.Recordset.Fields("M_ANO")) - 1 & "! Ignorando...", vbCritical
            End If
            
            If IsNull(ADOREG.Recordset.Fields("M_DT_DEM")) Then
                
                'COD da loja do cx do ANO atual
                Set w_ado_venda = de.cnc.Execute("SELECT TAB_VENDA.V_VR From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND TAB_FUNCIONARIO.F_DT_DEM IS NULL AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & Format(ADOREG.Recordset.Fields("M_MES"), "00") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD")).Clone
                If Not w_ado_venda.EOF Then
                    vrVenda = w_ado_venda.Fields(0)
                Else
                    MsgBox "Não há lançamentos do logo " & ADOREG.Recordset.Fields("M_LOGO") & " para o período: " & Format(ADOREG.Recordset.Fields("M_MES"), "00") & " / " & ADOREG.Recordset.Fields("M_ANO") & "! Ignorando...", vbCritical
                End If
      
                'vrVenda = de.cnc.Execute("SELECT TAB_VENDA.V_VR From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND TAB_FUNCIONARIO.F_DEM_OK = 0 AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & Format(adoReg.Recordset.Fields("M_MES"), "00") & " AND Right(TAB_VENDA.V_DATA,4) = " & adoReg.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & adoReg.Recordset.Fields("M_F_COD")).Fields(0)
            
                'If vrVenda <> "" Then
                If Not w_ado_vendaAnt.EOF And Not w_ado_venda.EOF Then
                    vrFixo = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_VR_FIXO From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & ADOREG.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD")).Fields(0)
                    vrMinimo = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_VR_MINIMO From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & ADOREG.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD")).Fields(0)
                    percComis = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_COMIS From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & ADOREG.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD")).Fields(0)
                    
                    '**** BONUS DE META
                    vrMeta = 999
                    'vrMeta = de.cnc.Execute("SELECT MT_VALOR FROM TAB_META INNER JOIN TAB_FUNCIONARIO ON MT_F_LOJA = TAB_FUNCIONARIO.F_LOJA WHERE MT_MES=" & adoReg.Recordset.Fields("M_MES") & " AND MT_ANO=" & adoReg.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_Codigo=" & W_ADO_FICHA.Fields("M_F_COD")).Fields(0)
                    If ADOREG.Recordset.Fields("M_LOGO") <> "XX" Then
                        vrMeta = de.cnc.Execute("SELECT TAB_META.MT_VALOR FROM TAB_META, Lojb010 INNER JOIN TAB_FUNCIONARIO ON Lojb010.COD_LOJ = TAB_FUNCIONARIO.F_Cod_L WHERE (((TAB_META.MT_MES)=" & ADOREG.Recordset.Fields("M_MES") & ") AND ((TAB_META.MT_ANO)=" & ADOREG.Recordset.Fields("M_ANO") & ") AND ((TAB_FUNCIONARIO.F_Codigo)=" & ADOREG.Recordset.Fields("M_F_COD") & ") AND ((Lojb010.NUM)=Right([tab_meta].[mt_f_loja],2)))").Fields(0)
                    End If
                    '****
                    
                             
                    vrSalario = vrFixo + ((vrVenda * 1000) * (percComis / 100))
                    vrComis = (vrVenda * 1000) * (percComis / 100)
                    percVenda = FormatNumber((100 - (vrVenda / vrVendaAnt * 100)) * -1, 2)
                    
                If percVenda > 0 Then
                    'Lojas em promoção ganham metade do bonus
                    If ADOREG.Recordset.Fields("M_LOGO") = "AC" Or ADOREG.Recordset.Fields("M_LOGO") = "IJ" Or ADOREG.Recordset.Fields("M_LOGO") = "FH" Then
                        percVenda = percVenda / 2
                    End If
    
                    'If percVenda >= 40 Then percVenda = percVenda / 2
                    If ADOREG.Recordset.Fields("M_MES") = "12" Then
                        'vrBonus = percVenda * 10
                        vrBonus = 0
                    Else
                        vrBonus = percVenda * 30
                    End If
                    
                    'Algumas gerentes em algumas lojas não ganham bonus
                    If (ADOREG.Recordset.Fields("M_LOGO") = "SV" And ADOREG.Recordset.Fields("M_NOME") Like "JANAINA%") _
                        Or (ADOREG.Recordset.Fields("M_LOGO") = "LP") _
                        Or (ADOREG.Recordset.Fields("M_LOGO") = "RR") Then
                            vrBonus = 0
                    End If
                    
                Else
                    vrBonus = 0
                End If
    
                    
                'If vrBonus > 1000 And adoReg.Recordset.Fields("M_MES") <> "12" Then vrBonus = 1000
                If vrBonus > 500 Then vrBonus = 500
    
                    vrSalarioBonus = vrSalario + vrBonus
                    
                    If vrSalario < vrMinimo Then
                        vrSalarioBonus = vrMinimo + vrBonus
                    End If
                
                 '**** BONUS DE META
                 vrMetaBonus = 0
                 
                'DEZEMBRO
                If ADOREG.Recordset.Fields("m_mes") <= 12 Then
                    wPercMeta = 0
                    
                    If vrVenda >= (vrMeta * 1.1) Then
                        vrMetaBonus = vrSalario * 0.1
                        wPercMeta = 10
                    ElseIf vrVenda >= (vrMeta * 1.08) Then
                        vrMetaBonus = vrSalario * 0.08
                        wPercMeta = 8
                    ElseIf vrVenda >= (vrMeta * 1.05) Then
                        vrMetaBonus = vrSalario * 0.05
                        wPercMeta = 5
                    End If
                    If vrVenda >= (vrMeta * 1.05) Then
                       descMeta = " + Bonus Meta(" & Format(vrSalario, "0.00") & " * " & Format(wPercMeta, "0.0") & "% = " & Format(vrMetaBonus, "0.00") & ") "
                       vrSalarioBonus = vrSalarioBonus + vrMetaBonus
                    Else
                       descMeta = ""
                    End If
                    '****
                Else
                    If vrVenda >= vrMeta Then
                       'Bonus de 15% do Salario
                       vrMetaBonus = vrSalario * 0.15
                       descMeta = " + Bonus Meta(" & Format(vrSalario, "0.00") & " * 15% = " & Format(vrMetaBonus, "0.00") & ") "
                       vrSalarioBonus = vrSalarioBonus + vrMetaBonus
                    Else
                       descMeta = ""
                    End If
                    '****
                End If
                
                If vrSalario < vrMinimo Then
                    wDesc = "**NÃO ATINGIU O MÍNIMO** (" & Format(vrVenda, "0.00") & " x " & percComis & "% = " & Format(vrComis, "0.00") & ") + " & Format(vrFixo, "0.00") & " = " & Format(vrSalario, "0.00") & " < Mínimo(" & Format(vrMinimo, "0.00") & ") + Bônus(" & percVenda & "%=" & Format(vrBonus, "0.00") & ") " & descMeta & " = " & Format(vrSalarioBonus, "0.00")
                Else
                    wDesc = "(" & Format(vrVenda, "0.00") & " x " & percComis & "% = " & Format(vrComis, "0.00") & ") + " & Format(vrFixo, "0.00") & " = " & Format(vrSalario, "0.00") & " + Bônus(" & percVenda & "%=" & Format(vrBonus, "0.00") & ") " & descMeta & " = " & Format(vrSalarioBonus, "0.00")
                End If
                
                de.cnc.Execute ("DELETE FROM TAB_DESC_CALC Where (C_TP_CONTA = 20) And (C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFicha") & ")")
                de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFicha"), 20, "+", vrSalarioBonus, wDesc, "", "0", "0", "0", "0"
                
                
                End If
            End If
            'Dados Contas
            Lancamentos
            
            
        Case "X": 'CAIXA EXTRA
        
            'Dim dtIni, dtFim As Date 'Ja declarado em cima (no VENDEDOR)
 
            'Dim w_Dt, w_dtUlt As Date 'Ja declarado cm cima (no CAIXA)
            'Dim w_DtDiff, w_ultDiaMes As Integer 'Ja declarado cm cima (no CAIXA)
            'Dim w_Piso, w_Comis, w_Premio, w_PisoOriginal As Double 'Ja declarado cm cima (no CAIXA)
        
            w_Piso = 0
            w_Piso = de.cnc.Execute("SELECT F_VR_FIXO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
            If w_Piso = 0 Then
                w_Piso = de.cnc.Execute("SELECT F_VPISO_R FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
                If w_Piso = 0 Then
                    w_Piso = de.cnc.Execute("SELECT F_VPISO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
                End If
            End If
            
            
            de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & " AND (C_TP_CONTA = 22)")
            
            'w_ultDiaMes = 30
            w_ultDiaMes = Day(UltDiaMes(ADOREG.Recordset.Fields("m_mes"), ADOREG.Recordset.Fields("m_ano")))
            If w_ultDiaMes < 30 Then w_ultDiaMes = 30
         
            If ADOREG.Recordset.Fields("m_dt_reg") = "" Or IsNull(ADOREG.Recordset.Fields("m_dt_reg")) Then
                w_Dt = CVDate(ADOREG.Recordset.Fields("m_dt_adm"))
            Else
                w_Dt = CVDate(ADOREG.Recordset.Fields("m_dt_reg"))
            End If
             
             
            If IsDate(ADOREG.Recordset.Fields("M_DT_DEM")) Then
                w_dtUlt = CVDate(ADOREG.Recordset.Fields("M_DT_DEM"))
            Else
                w_dtUlt = CVDate(UltDiaMes(ADOREG.Recordset.Fields("m_mes"), ADOREG.Recordset.Fields("m_ano")))
                'If Day(w_dtUlt) = 31 Then w_dtUlt = w_dtUlt - 1
                If Day(w_dtUlt) = 28 Then w_dtUlt = w_dtUlt + 2
                If Day(w_dtUlt) = 29 Then w_dtUlt = w_dtUlt + 1
            End If
             
            If Month(w_Dt) < Month(w_dtUlt) Then w_Dt = CVDate("01/" & Month(w_dtUlt) & "/" & Year(w_dtUlt))
             
            w_DtDiff = DateDiff("d", w_Dt, w_dtUlt) + 1
             
            w_PisoOriginal = w_Piso
             'MsgBox "Diff: " & w_DtDiff & " - Ini: " & w_Dt & " - Final: " & w_dtUlt
            If (w_DtDiff < w_ultDiaMes) And (w_DtDiff < 30) Then
                If w_ultDiaMes < 30 Then
                    w_Piso = ((w_Piso / w_ultDiaMes) * w_DtDiff)
                Else
                    w_Piso = ((w_Piso / 30) * w_DtDiff)
                    w_ultDiaMes = 30
                End If
                 
                de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 22, "+", Format(w_Piso, "0.00"), "#FIXO PROPORCIONAL# Fixo: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00") & " [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")

            Else
                de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 22, "+", Format(w_Piso, "0.00"), "#FIXO# Fixo: " & Format(w_Piso, "0.00") & " [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
            End If
     
            Lancamentos
            
        Case "2": '2o CAIXA
        
            'Dim dtIni, dtFim As Date 'Ja declarado em cima (no VENDEDOR)
 
            'Dim w_Dt, w_dtUlt As Date 'Ja declarado cm cima (no CAIXA)
            'Dim w_DtDiff, w_ultDiaMes As Integer 'Ja declarado cm cima (no CAIXA)
            'Dim w_Piso, w_Comis, w_Premio, w_PisoOriginal As Double 'Ja declarado cm cima (no CAIXA)
        
            w_Piso = 0
            w_Piso = de.cnc.Execute("SELECT F_VR_FIXO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
            If w_Piso = 0 Then
                w_Piso = de.cnc.Execute("SELECT F_VPISO_R FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
                If w_Piso = 0 Then
                    w_Piso = de.cnc.Execute("SELECT F_VPISO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
                End If
            End If
            
            
            de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & " AND (C_TP_CONTA = 22)")
            
            'w_ultDiaMes = 30
            w_ultDiaMes = Day(UltDiaMes(ADOREG.Recordset.Fields("m_mes"), ADOREG.Recordset.Fields("m_ano")))
            If w_ultDiaMes < 30 Then w_ultDiaMes = 30
         
            If ADOREG.Recordset.Fields("m_dt_reg") = "" Or IsNull(ADOREG.Recordset.Fields("m_dt_reg")) Then
                w_Dt = CVDate(ADOREG.Recordset.Fields("m_dt_adm"))
            Else
                w_Dt = CVDate(ADOREG.Recordset.Fields("m_dt_reg"))
            End If
             
             
            If IsDate(ADOREG.Recordset.Fields("M_DT_DEM")) Then
                w_dtUlt = CVDate(ADOREG.Recordset.Fields("M_DT_DEM"))
            Else
                w_dtUlt = CVDate(UltDiaMes(ADOREG.Recordset.Fields("m_mes"), ADOREG.Recordset.Fields("m_ano")))
                'If Day(w_dtUlt) = 31 Then w_dtUlt = w_dtUlt - 1
                If Day(w_dtUlt) = 28 Then w_dtUlt = w_dtUlt + 2
                If Day(w_dtUlt) = 29 Then w_dtUlt = w_dtUlt + 1
            End If
             
            If Month(w_Dt) < Month(w_dtUlt) Then w_Dt = CVDate("01/" & Month(w_dtUlt) & "/" & Year(w_dtUlt))
             
            w_DtDiff = DateDiff("d", w_Dt, w_dtUlt) + 1
             
            w_PisoOriginal = w_Piso
             'MsgBox "Diff: " & w_DtDiff & " - Ini: " & w_Dt & " - Final: " & w_dtUlt
            If (w_DtDiff < w_ultDiaMes) And (w_DtDiff < 30) Then
                If w_ultDiaMes < 30 Then
                    w_Piso = ((w_Piso / w_ultDiaMes) * w_DtDiff)
                Else
                    w_Piso = ((w_Piso / 30) * w_DtDiff)
                    w_ultDiaMes = 30
                End If
                 
                de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 22, "+", Format(w_Piso, "0.00"), "#FIXO PROPORCIONAL# Fixo: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00") & " [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")

            Else
                de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 22, "+", Format(w_Piso, "0.00"), "#FIXO# Fixo: " & Format(w_Piso, "0.00") & " [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
            End If
     
            Lancamentos
        
        
        Case "C": 'CAIXA
        
        'Dim vrMeta, vrMetaBonus, descMeta
        
            de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & " AND (C_TP_CONTA = 22)")
             'If de.rscmdTotalVend.State = 1 Then de.rscmdTotalVend.Close
             'de.cmdTotalVend TXT_MES, TXT_ANO, W_ADO_FICHA.Fields("M_LOGO")
             
             '*** looping entre os Vendedores p/ Calc. Média
             'W_QT = 1
             'W_TT = 0
             'w_DESCR = ""
             'Do While Not de.rscmdTotalVend.EOF
             '    W_TT = W_TT + de.rscmdTotalVend.Fields("valor")
             '    w_DESCR = w_DESCR & IIf(w_DESCR = "", "< (" & Format(de.rscmdTotalVend.Fields("valor"), "0.00"), " + " & Format(de.rscmdTotalVend.Fields("valor"), "0.00"))
             '
             '    If W_QT = IIf(IsNull(W_ADO_FICHA.Fields("CX_QT_VND")), 3, W_ADO_FICHA.Fields("CX_QT_VND")) Then
             '        w_Media = W_TT / W_QT
             '        w_DESCR = w_DESCR & ") = " & Format(W_TT, "0.00") & " / " & W_QT & " = " & Format(w_Media, "0.00") & " >"
             '        Exit Do
             '    End If
             '    W_QT = W_QT + 1
             '    de.rscmdTotalVend.MoveNext
             'Loop
        'txt_notas.Text = ("SELECT TAB_VENDA.V_VR From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND TAB_FUNCIONARIO.F_DEM_OK = 0 AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & Format(adoReg.Recordset.Fields("M_MES"), "00") & " AND Right(TAB_VENDA.V_DATA,4) = " & CInt(adoReg.Recordset.Fields("M_ANO")) - 1 & " AND TAB_FUNCIONARIO.F_CODIGO = " & W_ADO_FICHA.Fields("M_F_COD"))
        'Exit Sub
        
        'COD da loja do cx do ANO anterior
        Set w_ado_vendaAnt = de.cnc.Execute("SELECT TAB_VENDA.V_VR From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND TAB_FUNCIONARIO.F_DT_DEM IS NULL AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & Format(ADOREG.Recordset.Fields("M_MES"), "00") & " AND Right(TAB_VENDA.V_DATA,4) = " & CInt(ADOREG.Recordset.Fields("M_ANO")) - 1 & " AND TAB_FUNCIONARIO.F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD")).Clone
        If Not w_ado_vendaAnt.EOF Then
            vrVendaAnt = w_ado_vendaAnt.Fields(0)
        Else
            MsgBox "Não há lançamentos do logo " & ADOREG.Recordset.Fields("M_LOGO") & " para o período: " & Format(ADOREG.Recordset.Fields("M_MES"), "00") & " / " & CInt(ADOREG.Recordset.Fields("M_ANO")) - 1 & "! Ignorando...", vbCritical
        End If
        
        'COD da loja do cx do ANO atual
        Set w_ado_venda = de.cnc.Execute("SELECT TAB_VENDA.V_VR From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND TAB_FUNCIONARIO.F_DT_DEM IS NULL AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & Format(ADOREG.Recordset.Fields("M_MES"), "00") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD")).Clone
        If Not w_ado_venda.EOF Then
            vrVenda = w_ado_venda.Fields(0)
        Else
            vrVenda = 0
            MsgBox "Não há lançamentos do logo " & ADOREG.Recordset.Fields("M_LOGO") & " para o período: " & Format(ADOREG.Recordset.Fields("M_MES"), "00") & " / " & ADOREG.Recordset.Fields("M_ANO") & "! Ignorando...", vbCritical
        End If
    
        If Not w_ado_vendaAnt.EOF And Not w_ado_venda.EOF Then
            vrFixo = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_VR_FIXO From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & ADOREG.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD")).Fields(0)
            vrMinimo = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_VR_MINIMO From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & ADOREG.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD")).Fields(0)
            'perc1 = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_COMIS From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & ADOREG.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & adoReg.Recordset.Fields("M_F_COD")).Fields(0)
            'perc2 = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_COMIS2 From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & ADOREG.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & adoReg.Recordset.Fields("M_F_COD")).Fields(0)
            'perc3 = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_COMIS3 From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & ADOREG.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & adoReg.Recordset.Fields("M_F_COD")).Fields(0)
            If ADOREG.Recordset.Fields("M_MES") = 12 Then
                percDez = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_PERC_FIXO_DEZ From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & ADOREG.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD")).Fields(0)
                vrDez = vrFixo * percDez
                descDez = " + " & (percDez * 100) & "% (" & vrDez & ") "
            Else
                percDez = 0
                vrDez = 0
                descDez = ""
            End If
            
            'percComis = de.cnc.Execuezte("SELECT TAB_FUNCIONARIO.F_COMIS From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & adoReg.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & adoReg.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & adoReg.Recordset.Fields("M_F_COD")).Fields(0)
            
            '**** BONUS DE META
             vrMeta = 999
             'vrMeta = de.cnc.Execute("SELECT MT_VALOR FROM TAB_META INNER JOIN TAB_FUNCIONARIO ON MT_F_LOJA = TAB_FUNCIONARIO.F_LOJA WHERE MT_MES=" & adoReg.Recordset.Fields("M_MES") & " AND MT_ANO=" & adoReg.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_Codigo=" & adoReg.Recordset.Fields("M_F_COD")).Fields(0)
            If ADOREG.Recordset.Fields("M_LOGO") <> "XX" And ADOREG.Recordset.Fields("M_LOGO") <> "99" Then
                vrMeta = de.cnc.Execute("SELECT TAB_META.MT_VALOR FROM TAB_META, Lojb010 INNER JOIN TAB_FUNCIONARIO ON Lojb010.COD_LOJ = TAB_FUNCIONARIO.F_Cod_L WHERE (((TAB_META.MT_MES)=" & ADOREG.Recordset.Fields("M_MES") & ") AND ((TAB_META.MT_ANO)=" & ADOREG.Recordset.Fields("M_ANO") & ") AND ((TAB_FUNCIONARIO.F_Codigo)=" & ADOREG.Recordset.Fields("M_F_COD") & ") AND ((Lojb010.NUM)=Right([tab_meta].[mt_f_loja],2)))").Fields(0)
            End If
            '****
            
            'percVenda = (100 - (vrVenda / vrVendaAnt * 100)) * -1
            percVenda = (100 - (vrVenda / vrMeta * 100)) * -1
            
            'If percVenda < 5 Then
            '    percComis = perc1
            'ElseIf percVenda <= 10 Then
            '    percComis = perc2
            'ElseIf percVenda > 10 Then
            '    percComis = perc3
            'Else
            '    MsgBox "Pecentual sobre os lançamentos (" & percVenda & ") incorreto! Impossível continuar, cancelando!"
            '    Exit Sub
            'End If
            
            percComis = 0.45
            
            vrComis = (vrVenda * 1000) * (percComis / 100)
            vrSalario = vrFixo + (vrComis) + vrDez
            
            If vrSalario < vrMinimo Then
                vrSalario = vrMinimo
            End If
                                    
                                    
 
           
                                    
                                    
            '**** BONUS DE META
             vrMetaBonus = 0
             wFixoMeta = 0
             wPercMeta = 0
            If ADOREG.Recordset.Fields("M_MES") <= 12 Then
                
                If vrVenda >= (vrMeta) Then
                    wFixoMeta = 100
                    If (((vrVenda / vrMeta) - 1) * 100) >= 10 Then
                        vrMetaBonus = vrSalario * 0.1
                        wPercMeta = 10
                    Else
                        vrMetaBonus = vrSalario * ((vrVenda / vrMeta) - 1)
                        wPercMeta = ((vrVenda / vrMeta) - 1) * 100
                    End If
                    
                   descMeta = " + Bonus Meta(" & Format(vrSalario, "0.00") & " * " & Format(wPercMeta, "0.0") & "% = " & Format(vrMetaBonus, "0.00") & " + Fixo Meta(" & Format(wFixoMeta, "0.00") & ") = " & Format(vrMetaBonus + wFixoMeta, "0.00") & ") "
                   vrSalario = vrSalario + vrMetaBonus + wFixoMeta
                Else
                   descMeta = ""
                End If
            End If
            '*****

            If vrSalario < vrMinimo Then
                wDesc = "**NÃO ATINGIU O MÍNIMO** " & Format(vrVenda, "0.00") & " * " & Format(percComis, "0.00") & "% = " & Format(vrComis, "0.00") & " + " & Format(vrFixo, "0.00") & descDez & descMeta & " = " & Format(vrSalario, "0.00")
            Else
                wDesc = Format(vrVenda, "0.00") & " * " & Format(percComis, "0.00") & "% = " & Format(vrComis, "0.00") & " + " & Format(vrFixo, "0.00") & descDez & descMeta & " = " & Format(vrSalario, "0.00")
            End If
                                           
             '*** Pega o Piso referente se for com ou sem registro
             If IsNull(ADOREG.Recordset.Fields("m_Dt_Reg")) Then
                 w_Piso = ADOREG.Recordset.Fields("F_vpiso")
                 w_Pdesc = "Ps. B"
             Else
                 w_Piso = ADOREG.Recordset.Fields("F_vpiso_R")
                 w_Pdesc = "Ps. L"
             End If
             w_Piso = IIf(IsNull(w_Piso), 0, w_Piso)
             
             '*** paga comissão *** da média
             If vrSalario >= w_Piso Then
                 'w_desc = "CX - " & w_Pdesc & " : " & IIf(IsNull(w_Piso), "R$ 0,00", Format(w_Piso, "R$ 0.00")) & "   " & w_DESCR
                 de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), "22", "+", vrSalario, wDesc, 0, 0, 0, 0, 0
                 W_REG_CX = W_REG_CX + 1
             '*** paga piso ***
             Else
             
                    W_DT_INI_MES = CVDate("01/" & TXT_MES & "/" & TXT_ANO)
                    W_DT_FIM_MES = CVDate("01/" & Format(W_DT_INI_MES + 35, "MM/YYYY"))
                    'sE DT DE ADM. FOR MAIOR Q/ A DT DO PRIMEIRO DIA DO MES ***
                    If CVDate(ADOREG.Recordset.Fields("M_DT_ADM")) >= CVDate(W_DT_INI_MES) Then
                         W_DT_INI_MES = CVDate(ADOREG.Recordset.Fields("M_DT_ADM"))
                    End If
                    
                    If Not IsNull(ADOREG.Recordset.Fields("M_DT_DEM")) Then
                          W_QT_DIAS_TRAB = (CVDate(ADOREG.Recordset.Fields("M_DT_DEM")) + 1) - CVDate(W_DT_INI_MES)
                    ElseIf W_DT_INI_MES = CVDate("01/" & TXT_MES & "/" & TXT_ANO) Then
                          W_QT_DIAS_TRAB = "-30"
                    Else
                          W_QT_DIAS_TRAB = W_DT_FIM_MES - W_DT_INI_MES
                          
                    End If
                    
                    
                    '*** INCLUI PISO S/ REGISTRO ***
                    If IsNull(ADOREG.Recordset.Fields("M_DT_REG")) Then
                        If W_QT_DIAS_TRAB = "-30" Then
                            W_VALOR_PISO = ADOREG.Recordset.Fields("F_vpiso")
                            w_desc = "CX - " & w_Pdesc & " : " & IIf(IsNull(W_VALOR_PISO), "R$ 0,00", Format(W_VALOR_PISO, "R$ 0.00")) & "   " & w_DESCR
                        Else
                            W_VALOR_PISO = W_QT_DIAS_TRAB * (ADOREG.Recordset.Fields("F_vpiso") / 30)
                            w_desc = "CX - " & W_QT_DIAS_TRAB & " dias ref. ao " & w_Pdesc & " " & IIf(IsNull(W_VALOR_PISO), "R$ 0,00", Format(ADOREG.Recordset.Fields("F_vpiso"), "R$ 0.00")) & " :  (" & Format(ADOREG.Recordset.Fields("F_vpiso"), "R$ 0.00") & " / 30 = " & Format(ADOREG.Recordset.Fields("F_vpiso") / 30, "R$ 0.00") & " x " & W_QT_DIAS_TRAB & ")"
                        End If
                            
                        de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), "22", "+", CDbl(W_VALOR_PISO), w_desc, 0, 0, 0, 0, 0
                        W_REG_CX = W_REG_CX + 1
                        
                    '*** INCLUI PISO C/ REGISTRO ***
                    Else
                        If W_QT_DIAS_TRAB = "-30" Then
                            W_VALOR_PISO = ADOREG.Recordset.Fields("F_vpiso_R")
                            w_desc = "CX - " & w_Pdesc & " : " & IIf(IsNull(W_VALOR_PISO), "R$ 0,00", Format(W_VALOR_PISO, "R$ 0.00")) & "   " & w_DESCR
                        Else
                            W_VALOR_PISO = W_QT_DIAS_TRAB * (ADOREG.Recordset.Fields("F_vpiso_R") / 30)
                            w_desc = "CX - " & W_QT_DIAS_TRAB & " dias ref. ao " & w_Pdesc & " " & IIf(IsNull(W_VALOR_PISO), "R$ 0,00", Format(ADOREG.Recordset.Fields("F_vpiso_R"), "R$ 0.00")) & " :  (" & Format(ADOREG.Recordset.Fields("F_vpiso_R"), "R$ 0.00") & " / 30) = " & Format(ADOREG.Recordset.Fields("F_vpiso_R") / 30, "R$ 0.00") & " x " & W_QT_DIAS_TRAB & ")"
                        End If
                        
                        If IsNull(W_VALOR_PISO) Then W_VALOR_PISO = 0
                        
                        de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), "22", "+", CDbl(W_VALOR_PISO), w_desc, 0, 0, 0, 0, 0
                        
                        W_REG_CX = W_REG_CX + 1
                    End If
                End If
             End If

        Lancamentos
        End Select
        
End Sub

Private Sub cmdAddSalarioGerente_Click()

End Sub

Private Sub cmdATTotal_Click()
On Error GoTo err1

        'Atualização dos Totais das fichas
        If de.rsTAB_FICHA_MENS.State = 1 Then de.rsTAB_FICHA_MENS.Close
        de.TAB_FICHA_MENS
        'FILTRA AS FICHAS DO MÊS E Q/ NÃO ESTA BLOQUEADA
        de.rsTAB_FICHA_MENS.Filter = "M_MES = " & TXT_MES & " and M_Ano = " & TXT_ANO & ""
        de.rsTAB_FICHA_MENS.MoveFirst
        On Error Resume Next
        Do While Not de.rsTAB_FICHA_MENS.EOF
        
            W_MAIS = 0
            W_MENOS = 0
            W_TOTAL = 0
            '*** CALCULA O TOTAL - APÓS O NOVO VALOR ***
            W_MAIS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MAIS FROM TAB_DESC_CALC  WHERE (C_TP_OP = '+') AND (C_N_FICHA = " & de.rsTAB_FICHA_MENS.Fields("M_NFICHA") & ")").Fields("MAIS")
            W_MENOS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MENOS FROM TAB_DESC_CALC WHERE (C_TP_OP = '-') AND (C_N_FICHA = " & de.rsTAB_FICHA_MENS.Fields("M_NFICHA") & ")").Fields("MENOS")
            W_TOTAL = IIf(IsNull(W_MENOS), 0, W_MENOS) + IIf(IsNull(W_MAIS), 0, W_MAIS)
            '*** Atualiza os Campos  Total , Mais e Menos
            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_TOTAL = '" & CDbl(IIf(IsNull(W_TOTAL), 0, W_TOTAL)) & "', M_TOTAL_MAIS = '" & CDbl(IIf(IsNull(W_MAIS), 0, W_MAIS)) & "', M_TOTAL_MENOS = '" & CDbl(IIf(IsNull(W_MENOS), 0, W_MENOS)) & "' WHERE (M_NFICHA = " & de.rsTAB_FICHA_MENS.Fields("M_NFICHA") & ")"
        
        
        de.rsTAB_FICHA_MENS.MoveNext
        Loop
        
        Cancelar

        MsgBox "Atualização dos Totais feita com sucesso!", vbInformation

sair:
    Exit Sub
err1:
    
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    
    Resume sair
End Sub



Sub cmdAtualizar_Click()
    frmAtualizando.Visible = True
    Pause 1
    Form_Load
    frmAtualizando.Visible = False
End Sub

Private Sub cmdAtualizarCaption_Click()
 cmdAtualizar_Click
End Sub

Private Sub cmdBonus_Click()
Dim dtIni, dtFim As Date

If MsgBox("Gerar Custos de Camisetas BÔNUS p/ o mês " & lblMes.Caption & "?", vbYesNo, "CUSTOS CAMISETAS BÔNUS") = vbYes Then
    
    frm_ESCOLHA_DATA.Show 1
    
    dtIni = CDate(frm_ESCOLHA_DATA.TXT_DT_INICIAL)
    dtFim = CDate(frm_ESCOLHA_DATA.TXT_DT_FINAL)
   
    If de.rscmdBonus.State = 1 Then de.rscmdBonus.Close
    
    On Error Resume Next
    'de.cmdDROPtmpBonus
    'de.cmdCREATEtmpBonus
    de.cmdDELETEtmpBonus
    de.cmdAddtmpBonus dtIni, dtFim
    
    de.cmdBonus
    
    
    Dim wNumContas
        wNumContas = 99
    
        'CUSTOS CAMISETAS BONUS
        Dim custoUN, custoTOT As Double
        Dim w_descricao As String
        Dim qtBONUS As Integer
        qtBONUS = 0
        
        Dim adoBONUS As ADODB.Recordset
        Set adoBONUS = de.rscmdBonus.Clone

        If de.cncContas.State = 0 Then de.cncContas.Open
        Do While Not adoBONUS.EOF
            
            wNumContas = de.cncContas.Execute("SELECT tblloja_loja_contas FROM tblloja where tblloja_loja_fichas = '" & Format(adoBONUS.Fields("mov_estoque_filial_filial_nome"), "000") & "'").Fields(0)
            If wNumContas <> 99 Then
                custoUN = 0
                Select Case adoBONUS.Fields("COD_MARCA")
                    'ALLDRESS
                    Case "00003":
                        custoUN = 10
                    'HOT LIFE/SPORT CO/OAHU
                    Case "00862":
                        custoUN = 6.8
                    'FEDERAL TRADE
                    Case "00033":
                        custoUN = 7.5
                    'RADICAL
                    Case "00105":
                        custoUN = 6.8
                End Select
                
                If custoUN = 0 Then custoUN = 6.8
                
                custoTOT = custoUN * (adoBONUS.Fields("sum_mov_estoque_quantidade_") * -1)
                w_descricao = "**Incluído automaticamente** Marca: " & adoBONUS.Fields("F_615818150") & " | Qtde: " & (adoBONUS.Fields("sum_mov_estoque_quantidade_") * -1) & " | Custo: " & custoUN
                de.cmdAddProcesso "1622", CDate(Now()), w_descricao, "R", CStr(custoTOT), wNumContas, CDate(Now())
                lastNumProcesso = de.cncContas.Execute("SELECT MAX(tblpro_num) FROM tblpropg").Fields(0)
                de.cmdAddProcessoItem lastNumProcesso, CDate("01/" & TXT_MES & "/" & TXT_ANO), CStr(custoTOT), CDate("01/" & TXT_MES & "/" & TXT_ANO)
                qtBONUS = qtBONUS + 1
            End If
            adoBONUS.MoveNext

        Loop
        MsgBox "Foram incluídos " & qtBONUS & " processos de CUSTOS DE CAMISETAS BONUS no sistema de Contas, na data de " & CDate("01/" & TXT_MES & "/" & TXT_ANO) & ".", vbInformation, "Exportação de Gasolina"
        'FIM CUSTOS CAMISETAS BONUS

End If

End Sub

Private Sub cmdComerciario_Click()
'On Error GoTo err1
        Dim w_DiasTrabalhados, valorComerciario, diasComerciario
        Dim adoFuncs As ADODB.Recordset
        
        Set adoFuncs = de.cnc.Execute("SELECT DISTINCT (TAB_FICHA_MENS.M_NFICHA) AS FICHA, TAB_FICHA_MENS.M_DT_ADM, TAB_FICHA_MENS.M_DT_DEM, " & _
                                " TAB_FICHA_MENS.M_F_COD AS COD, TAB_FICHA_MENS.M_LOGO, TAB_FUNCIONARIO.F_VPISO_R " & _
                                " FROM TAB_FICHA_MENS INNER JOIN TAB_FUNCIONARIO " & _
                                " ON TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo " & _
                                " WHERE (((TAB_FICHA_MENS.M_ANO)=" & TXT_ANO & ") AND ((TAB_FICHA_MENS.M_MES)=" & TXT_MES & "))").Clone
                                
    
        Do While Not adoFuncs.EOF
          
                If adoFuncs.Fields("M_LOGO") <> "RP" Then
                    de.cnc.Execute ("DELETE FROM TAB_DESC_CALC Where (C_TP_CONTA = 82) And (C_N_FICHA = " & adoFuncs.Fields("FICHA") & ")")
                    
                    
                    If IsDate(adoFuncs.Fields("M_DT_DEM")) Then
                        w_DiasTrabalhados = DateDiff("d", adoFuncs.Fields("M_DT_ADM"), adoFuncs.Fields("M_DT_DEM"))
                    Else
                        w_DiasTrabalhados = DateDiff("d", adoFuncs.Fields("M_DT_ADM"), Date)
                    End If
                    
                    If IsNull(adoFuncs.Fields("F_VPISO_R")) Then
                        valorComerciario = 0
                    Else
                    
                        valorComerciario = adoFuncs.Fields("F_VPISO_R") / 30
                        If w_DiasTrabalhados <= 90 Then
                            diasComerciario = 0
                        ElseIf w_DiasTrabalhados <= 180 Then
                            diasComerciario = 1
                        Else
                            If adoFuncs.Fields("M_LOGO") = "RR" Or adoFuncs.Fields("M_LOGO") = "RH" Or adoFuncs.Fields("M_LOGO") = "RA" Or adoFuncs.Fields("M_LOGO") = "AV" Then
                                diasComerciario = 1
                            Else
                                diasComerciario = 2
                            End If
                        End If
                        
                        valorComerciario = valorComerciario * diasComerciario
                    End If
                    
                    de.cmdIncluirDescCalc2 Date, adoFuncs.Fields("FICHA"), 82, "+", valorComerciario, "Dia do Comerciário referente à " & diasComerciario & " dia(s)", "0", "0", "0", "0", adoFuncs.Fields("COD"), 0
                End If
              
            adoFuncs.MoveNext
        Loop

    ADOREG.Recordset.MoveFirst
        
    'Dados Contas
    Lancamentos
    
    
sair:
    Exit Sub
err1:
    
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    
    Resume sair
End Sub

Private Sub cmdComisCx_Click()
    Dim vrVenda, vrVendaAnt, vrFixo, vrMinimo, percComis, vrSalario, vrComis, percVenda, perc1, perc2, perc3, percDez, vrDez, descDez, vrMeta, vrMetaBonus, descMeta, wFixoMeta
    Dim W_ADO_FICHA As ADODB.Recordset

    cbMostrar.text = "CAIXA"
    cmdMostrar_Click
    
    'If de.rscmdSqlComissao.State = 1 Then de.rscmdSqlComissao.Close
    'de.cmdSqlComissao TXT_LOGO, Format(TXT_MES, "00"), TXT_ANO
     
     '*** SQL de Premio ***
     'Set w_ado_Premio = de.cnc.Execute("SELECT P_LOJA, SUM(P_VALOR_PG) AS premio, P_VENDEDOR FROM TAB_temp WHERE P_ORDEM > 0 GROUP BY P_LOJA, P_VENDEDOR").Clone
     
     '*** ABRE AS FICHAS ***
    '                    Set W_ADO_FICHA = de.cnc.Execute("SELECT TAB_FICHA_MENS.M_MES, TAB_FICHA_MENS.M_ANO, TAB_FICHA_MENS.M_NFICHA, TAB_FICHA_MENS.M_LOGO, MID(TAB_FUNCIONARIO.F_COD_CENTRAL, 3) AS COD_FUNC_CENTRAL, TAB_FUNCIONARIO.F_TIPO as TIPO ,  TAB_FUNCIONARIO.F_VPISO as VPISO, TAB_FUNCIONARIO.F_VPISO_R as VPISO_R, TAB_FUNCIONARIO.F_CX_QT_VND as CX_QT_VND, TAB_FICHA_MENS.M_DT_REG AS DT_REG, TAB_FICHA_MENS.M_DT_DEM AS DT_DEM, TAB_FICHA_MENS.M_DT_ADM AS DT_ADM, TAB_FICHA_MENS.M_F_COD   FROM TAB_FICHA_MENS, TAB_FUNCIONARIO WHERE TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_LOGO LIKE '" & UCase(TXT_LOGO) & "' AND TAB_FICHA_MENS.M_ACORDO = 0 AND (TAB_FICHA_MENS.M_BLOQ = 0)) AND (M_COMISSAO = 'N') ORDER BY TAB_FICHA_MENS.M_LOGO, MID(TAB_FUNCIONARIO.F_COD_CENTRAL, 3)").Clone
    'Set W_ADO_FICHA = de.cnc.Execute("SELECT TAB_FICHA_MENS.M_MES, TAB_FICHA_MENS.M_ANO, TAB_FICHA_MENS.M_NFICHA, TAB_FICHA_MENS.M_LOGO, MID(TAB_FUNCIONARIO.F_COD_CENTRAL, 3) AS COD_FUNC_CENTRAL, TAB_FUNCIONARIO.F_TIPO as TIPO ,  TAB_FUNCIONARIO.F_VPISO as VPISO, TAB_FUNCIONARIO.F_VPISO_R as VPISO_R, TAB_FUNCIONARIO.F_CX_QT_VND as CX_QT_VND, TAB_FICHA_MENS.M_DT_REG AS DT_REG, TAB_FICHA_MENS.M_DT_DEM AS DT_DEM, TAB_FICHA_MENS.M_DT_ADM AS DT_ADM, TAB_FICHA_MENS.M_F_COD   FROM TAB_FICHA_MENS, TAB_FUNCIONARIO WHERE TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_LOGO LIKE '" & UCase(TXT_LOGO) & "' AND TAB_FICHA_MENS.M_ACORDO = 0 AND (TAB_FICHA_MENS.M_BLOQ = 0))  ORDER BY TAB_FICHA_MENS.M_LOGO, MID(TAB_FUNCIONARIO.F_COD_CENTRAL, 3)").Clone
    Set W_ADO_FICHA = de.cnc.Execute("SELECT TAB_FICHA_MENS.M_NOME, TAB_FICHA_MENS.M_MES, TAB_FICHA_MENS.M_ANO, TAB_FICHA_MENS.M_NFICHA, TAB_FICHA_MENS.M_LOGO, MID(TAB_FUNCIONARIO.F_COD_CENTRAL, 3) AS COD_FUNC_CENTRAL, TAB_FICHA_MENS.M_TIPO as TIPO ,  TAB_FUNCIONARIO.F_VPISO as VPISO, TAB_FUNCIONARIO.F_VPISO_R as VPISO_R, TAB_FUNCIONARIO.F_CX_QT_VND as CX_QT_VND, TAB_FICHA_MENS.M_DT_REG AS DT_REG, TAB_FICHA_MENS.M_DT_DEM AS DT_DEM, TAB_FICHA_MENS.M_DT_ADM AS DT_ADM, TAB_FICHA_MENS.M_F_COD FROM TAB_FICHA_MENS, TAB_FUNCIONARIO WHERE TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND TAB_FICHA_MENS.M_ACORDO = 0 AND (TAB_FICHA_MENS.M_DT_DEM IS NULL)  ORDER BY TAB_FICHA_MENS.M_LOGO, MID(TAB_FUNCIONARIO.F_COD_CENTRAL, 3)").Clone
    
    
    '***  entre os caixa **** calc a média ***
         'filtra as fichas somente dos caixas
         'W_ADO_FICHA.Filter = "TIPO = 'C' AND M_LOGO = 'FH'"
         W_ADO_FICHA.Filter = "TIPO = 'C'"
         'W_ADO_FICHA.Filter = "TIPO = 'C' OR TIPO = '2' "
         'If ck_Nome.value = 0 Then W_ADO_FICHA.Filter = "TIPO = 'C' AND M_F_COD = " & dbNome.BoundText & ""
         
        If Not W_ADO_FICHA.EOF Then W_ADO_FICHA.MoveFirst
         
         Do While Not W_ADO_FICHA.EOF
            de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & W_ADO_FICHA.Fields("M_NFICHA") & " AND (C_TP_CONTA = 22)")
             'If de.rscmdTotalVend.State = 1 Then de.rscmdTotalVend.Close
             'de.cmdTotalVend TXT_MES, TXT_ANO, W_ADO_FICHA.Fields("M_LOGO")
             
             '*** looping entre os Vendedores p/ Calc. Média
             'W_QT = 1
             'W_TT = 0
             'w_DESCR = ""
             'Do While Not de.rscmdTotalVend.EOF
             '    W_TT = W_TT + de.rscmdTotalVend.Fields("valor")
             '    w_DESCR = w_DESCR & IIf(w_DESCR = "", "< (" & Format(de.rscmdTotalVend.Fields("valor"), "0.00"), " + " & Format(de.rscmdTotalVend.Fields("valor"), "0.00"))
             '
             '    If W_QT = IIf(IsNull(W_ADO_FICHA.Fields("CX_QT_VND")), 3, W_ADO_FICHA.Fields("CX_QT_VND")) Then
             '        w_Media = W_TT / W_QT
             '        w_DESCR = w_DESCR & ") = " & Format(W_TT, "0.00") & " / " & W_QT & " = " & Format(w_Media, "0.00") & " >"
             '        Exit Do
             '    End If
             '    W_QT = W_QT + 1
             '    de.rscmdTotalVend.MoveNext
             'Loop
        'txt_notas.Text = ("SELECT TAB_VENDA.V_VR From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND TAB_FUNCIONARIO.F_DEM_OK = 0 AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & Format(adoReg.Recordset.Fields("M_MES"), "00") & " AND Right(TAB_VENDA.V_DATA,4) = " & CInt(adoReg.Recordset.Fields("M_ANO")) - 1 & " AND TAB_FUNCIONARIO.F_CODIGO = " & W_ADO_FICHA.Fields("M_F_COD"))
        'Exit Sub
        
        'COD da loja do cx do ANO anterior
        Set w_ado_vendaAnt = de.cnc.Execute("SELECT TAB_VENDA.V_VR From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND TAB_FUNCIONARIO.F_DT_DEM IS NULL AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & Format(ADOREG.Recordset.Fields("M_MES"), "00") & " AND Right(TAB_VENDA.V_DATA,4) = " & CInt(ADOREG.Recordset.Fields("M_ANO")) - 1 & " AND TAB_FUNCIONARIO.F_CODIGO = " & W_ADO_FICHA.Fields("M_F_COD")).Clone
        If Not w_ado_vendaAnt.EOF Then
            vrVendaAnt = w_ado_vendaAnt.Fields(0)
        Else
            MsgBox "Não há lançamentos do logo " & W_ADO_FICHA.Fields("M_LOGO") & " para o período: " & Format(ADOREG.Recordset.Fields("M_MES"), "00") & " / " & CInt(ADOREG.Recordset.Fields("M_ANO")) - 1 & "! Ignorando...", vbCritical
        End If
        
        'COD da loja do cx do ANO atual
        Set w_ado_venda = de.cnc.Execute("SELECT TAB_VENDA.V_VR From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND TAB_FUNCIONARIO.F_DT_DEM IS NULL AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & Format(ADOREG.Recordset.Fields("M_MES"), "00") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & W_ADO_FICHA.Fields("M_F_COD")).Clone
        If Not w_ado_venda.EOF Then
            vrVenda = w_ado_venda.Fields(0)
        Else
            vrVenda = 0
            MsgBox "Não há lançamentos do logo " & W_ADO_FICHA.Fields("M_LOGO") & " para o período: " & Format(ADOREG.Recordset.Fields("M_MES"), "00") & " / " & ADOREG.Recordset.Fields("M_ANO") & "! Ignorando...", vbCritical
        End If
    
        If Not w_ado_vendaAnt.EOF And Not w_ado_venda.EOF Then
            vrFixo = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_VR_FIXO From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & ADOREG.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & W_ADO_FICHA.Fields("M_F_COD")).Fields(0)
            vrMinimo = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_VR_MINIMO From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & ADOREG.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & W_ADO_FICHA.Fields("M_F_COD")).Fields(0)
            'perc1 = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_COMIS From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & ADOREG.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & W_ADO_FICHA.Fields("M_F_COD")).Fields(0)
            'perc2 = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_COMIS2 From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & ADOREG.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & W_ADO_FICHA.Fields("M_F_COD")).Fields(0)
            'perc3 = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_COMIS3 From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & ADOREG.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & W_ADO_FICHA.Fields("M_F_COD")).Fields(0)
            If ADOREG.Recordset.Fields("M_MES") = 12 Then
                percDez = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_PERC_FIXO_DEZ From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & ADOREG.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & W_ADO_FICHA.Fields("M_F_COD")).Fields(0)
                vrDez = vrFixo * percDez
                descDez = " + " & (percDez * 100) & "% (" & vrDez & ") "
            Else
                percDez = 0
                vrDez = 0
                descDez = ""
            End If
            
            'percComis = de.cnc.Execuezte("SELECT TAB_FUNCIONARIO.F_COMIS From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & adoReg.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & adoReg.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & W_ADO_FICHA.Fields("M_F_COD")).Fields(0)
            
            '**** BONUS DE META
             vrMeta = 999
             'vrMeta = de.cnc.Execute("SELECT MT_VALOR FROM TAB_META INNER JOIN TAB_FUNCIONARIO ON MT_F_LOJA = TAB_FUNCIONARIO.F_LOJA WHERE MT_MES=" & adoReg.Recordset.Fields("M_MES") & " AND MT_ANO=" & adoReg.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_Codigo=" & W_ADO_FICHA.Fields("M_F_COD")).Fields(0)
            If W_ADO_FICHA.Fields("M_LOGO") <> "XX" And W_ADO_FICHA.Fields("M_LOGO") <> "99" Then
                vrMeta = de.cnc.Execute("SELECT TAB_META.MT_VALOR FROM TAB_META, Lojb010 INNER JOIN TAB_FUNCIONARIO ON Lojb010.COD_LOJ = TAB_FUNCIONARIO.F_Cod_L WHERE (((TAB_META.MT_MES)=" & W_ADO_FICHA.Fields("M_MES") & ") AND ((TAB_META.MT_ANO)=" & W_ADO_FICHA.Fields("M_ANO") & ") AND ((TAB_FUNCIONARIO.F_Codigo)=" & W_ADO_FICHA.Fields("M_F_COD") & ") AND ((Lojb010.NUM)=Right([tab_meta].[mt_f_loja],2)))").Fields(0)
            End If
            '****
            
            'percVenda = (100 - (vrVenda / vrVendaAnt * 100)) * -1
            percVenda = (100 - (vrVenda / vrMeta * 100)) * -1
            
            'If percVenda < 5 Then
            '    percComis = perc1
            'ElseIf percVenda <= 10 Then
            '    percComis = perc2
            'ElseIf percVenda > 10 Then
            '    percComis = perc3
            'Else
            '    MsgBox "Pecentual sobre os lançamentos (" & percVenda & ") incorreto! Impossível continuar, cancelando!"
            '    Exit Sub
            'End If
            
            percComis = 0.45
            
            vrComis = (vrVenda * 1000) * (percComis / 100)
            vrSalario = vrFixo + (vrComis) + vrDez
            
              
            'If vrSalario < vrMinimo Then
            '    vrSalario = vrMinimo
            'End If
            
            '**** BONUS DE META
             vrMetaBonus = 0
             wFixoMeta = 0
             wPercMeta = 0
            If ADOREG.Recordset.Fields("M_MES") <= 12 Then
                
                If vrVenda >= (vrMeta) Then
                    wFixoMeta = 200
                    If (((vrVenda / vrMeta) - 1) * 100) >= 5 Then
                        vrMetaBonus = vrSalario * 0.05
                        wPercMeta = 5
                    Else
                        vrMetaBonus = vrSalario * ((vrVenda / vrMeta) - 1)
                        wPercMeta = ((vrVenda / vrMeta) - 1) * 100
                    End If
                   descMeta = " + Bonus Meta(" & Format(vrSalario, "0.00") & " * " & Format(wPercMeta, "0.0") & "% = " & Format(vrMetaBonus, "0.00") & " + Fixo Meta(" & Format(wFixoMeta, "0.00") & ") = " & Format(vrMetaBonus + wFixoMeta, "0.00") & ") "
                Else
                   descMeta = ""
                End If
                
                If vrSalario < vrMinimo Then
                     vrSalario = vrMinimo + vrMetaBonus + wFixoMeta
                Else
                     vrSalario = vrSalario + vrMetaBonus + wFixoMeta
                End If
                
            End If
            '*****


            If vrSalario < vrMinimo Then
                wDesc = "**NÃO ATINGIU O MÍNIMO** " & Format(vrVenda, "0.00") & " * " & Format(percComis, "0.00") & "% = " & Format(vrComis, "0.00") & " + " & Format(vrFixo, "0.00") & descDez & descMeta & " = " & Format(vrSalario, "0.00")
            Else
                wDesc = Format(vrVenda, "0.00") & " * " & Format(percComis, "0.00") & "% = " & Format(vrComis, "0.00") & " + " & Format(vrFixo, "0.00") & descDez & descMeta & " = " & Format(vrSalario, "0.00")
            End If
                                           
             '*** Pega o Piso referente se for com ou sem registro
             If IsNull(W_ADO_FICHA.Fields("Dt_Reg")) Then
                 w_Piso = W_ADO_FICHA.Fields("vpiso")
                 w_Pdesc = "Ps. B"
             Else
                 w_Piso = W_ADO_FICHA.Fields("vpiso_R")
                 w_Pdesc = "Ps. L"
             End If
             w_Piso = IIf(IsNull(w_Piso), 0, w_Piso)
             
             '*** paga comissão *** da média
             If vrSalario >= w_Piso Then
                 'w_desc = "CX - " & w_Pdesc & " : " & IIf(IsNull(w_Piso), "R$ 0,00", Format(w_Piso, "R$ 0.00")) & "   " & w_DESCR
                 de.cmdIncluirDescCalc Date, W_ADO_FICHA.Fields("M_NFICHA"), "22", "+", vrSalario, wDesc, 0, 0, 0, 0, 0
                 W_REG_CX = W_REG_CX + 1
             '*** paga piso ***
             Else
             
                    W_DT_INI_MES = CVDate("01/" & TXT_MES & "/" & TXT_ANO)
                    W_DT_FIM_MES = CVDate("01/" & Format(W_DT_INI_MES + 35, "MM/YYYY"))
                    'sE DT DE ADM. FOR MAIOR Q/ A DT DO PRIMEIRO DIA DO MES ***
                    If CVDate(W_ADO_FICHA.Fields("DT_ADM")) >= CVDate(W_DT_INI_MES) Then
                         W_DT_INI_MES = CVDate(W_ADO_FICHA.Fields("DT_ADM"))
                    End If
                    
                    If Not IsNull(W_ADO_FICHA.Fields("DT_DEM")) Then
                          W_QT_DIAS_TRAB = (CVDate(W_ADO_FICHA.Fields("DT_DEM")) + 1) - CVDate(W_DT_INI_MES)
                    ElseIf W_DT_INI_MES = CVDate("01/" & TXT_MES & "/" & TXT_ANO) Then
                          W_QT_DIAS_TRAB = "-30"
                    Else
                          W_QT_DIAS_TRAB = W_DT_FIM_MES - W_DT_INI_MES
                          
                    End If
                    
                    
                    '*** INCLUI PISO S/ REGISTRO ***
                    If IsNull(W_ADO_FICHA.Fields("DT_REG")) Then
                        If W_QT_DIAS_TRAB = "-30" Then
                            W_VALOR_PISO = W_ADO_FICHA.Fields("vpiso")
                            w_desc = "CX - " & w_Pdesc & " : " & IIf(IsNull(W_VALOR_PISO), "R$ 0,00", Format(W_VALOR_PISO, "R$ 0.00")) & "   " & w_DESCR
                        Else
                            W_VALOR_PISO = W_QT_DIAS_TRAB * (W_ADO_FICHA.Fields("vpiso") / 30)
                            w_desc = "CX - " & W_QT_DIAS_TRAB & " dias ref. ao " & w_Pdesc & " " & IIf(IsNull(W_VALOR_PISO), "R$ 0,00", Format(W_ADO_FICHA.Fields("vpiso"), "R$ 0.00")) & " :  (" & Format(W_ADO_FICHA.Fields("vpiso"), "R$ 0.00") & " / 30 = " & Format(W_ADO_FICHA.Fields("vpiso") / 30, "R$ 0.00") & " x " & W_QT_DIAS_TRAB & ")"
                        End If
                            
                        de.cmdIncluirDescCalc Date, W_ADO_FICHA.Fields("M_NFICHA"), "22", "+", CDbl(W_VALOR_PISO), w_desc, 0, 0, 0, 0, 0
                        W_REG_CX = W_REG_CX + 1
                        
                    '*** INCLUI PISO C/ REGISTRO ***
                    Else
                        If W_QT_DIAS_TRAB = "-30" Then
                            W_VALOR_PISO = W_ADO_FICHA.Fields("vpiso_R")
                            w_desc = "CX - " & w_Pdesc & " : " & IIf(IsNull(W_VALOR_PISO), "R$ 0,00", Format(W_VALOR_PISO, "R$ 0.00")) & "   " & w_DESCR
                        Else
                            W_VALOR_PISO = W_QT_DIAS_TRAB * (W_ADO_FICHA.Fields("vpiso_R") / 30)
                            w_desc = "CX - " & W_QT_DIAS_TRAB & " dias ref. ao " & w_Pdesc & " " & IIf(IsNull(W_VALOR_PISO), "R$ 0,00", Format(W_ADO_FICHA.Fields("vpiso_R"), "R$ 0.00")) & " :  (" & Format(W_ADO_FICHA.Fields("vpiso_R"), "R$ 0.00") & " / 30) = " & Format(W_ADO_FICHA.Fields("vpiso_R") / 30, "R$ 0.00") & " x " & W_QT_DIAS_TRAB & ")"
                        End If
                        
                        If IsNull(W_VALOR_PISO) Then W_VALOR_PISO = 0
                        
                        de.cmdIncluirDescCalc Date, W_ADO_FICHA.Fields("M_NFICHA"), "22", "+", CDbl(W_VALOR_PISO), w_desc, 0, 0, 0, 0, 0
                        
                        W_REG_CX = W_REG_CX + 1
                    End If
                End If
             End If
                                         
             W_ADO_FICHA.MoveNext
         Loop
    
End Sub

Private Sub cmdComisMwts_Click()
Dim dtIni, dtFim As Date
Dim adoComis As ADODB.Recordset
Dim adoComisTMP As ADODB.Recordset
Dim adoComisDump As ADODB.Recordset
Dim sqlComissTMP As String
Dim vrMeta, vrMetaBonus, descMeta, vrSalario, wFixoMeta
Dim w_ado_venda As ADODB.Recordset

    frm_ESCOLHA_DATA.Show 1
    
    dtIni = CDate(frm_ESCOLHA_DATA.TXT_DT_INICIAL)
    dtFim = CDate(frm_ESCOLHA_DATA.TXT_DT_FINAL)
   
    If de.rscmdComiss_Grouping.State = 1 Then de.rscmdComiss_Grouping.Close
    
    'On Error Resume Next
    'de.cmdDROPtmpComis1
    'de.cmdDROPtmpComis2
    
    'de.cmdCREATEtmpComis1
    'de.cmdCREATEtmpComis2
    
    de.cmdDELETEtmpComis1
    de.cmdDELETEtmpComis2
    
    de.cmdAddtmpComis1 dtIni, dtFim, dtIni, dtFim, dtIni, dtFim, dtIni, dtFim, dtIni, dtFim
    de.cmdAddtmpComis2 dtIni, dtFim, dtIni, dtFim
        
    de.cmdComiss_Grouping
    
    If MsgBox("Tem certeza que deseja (RE)GERAR A COMISSÃO DE VENDEDORES para " & lblMes.Caption & "?", vbYesNo, "GERAR COMISSÃO") = vbYes Then
    
        de.cmdComissGerar
        Set adoComis = de.rscmdComissGerar.Clone
            
        cbMostrar.text = "VENDEDOR"
        cmdMostrar_Click
            
        Dim w_Dt, w_dtUlt As Date
        Dim w_DtDiff, w_ultDiaMes As Integer
        Dim w_Piso, w_Comis, w_Premio, w_PisoOriginal, w_FixoOriginal, w_Fixo As Double
            
        ADOREG.Recordset.MoveFirst
        Do While Not ADOREG.Recordset.EOF
        
           If ADOREG.Recordset.Fields("M_TIPO") = "V" Then
        
            'If adoReg.Recordset.Fields("M_TIPO") = "V" And ( _
            '        adoReg.Recordset.Fields("M_LOGO") = "FH") Then
            '        adoReg.Recordset.Fields("M_LOGO") = "BA" Or _
            '        adoReg.Recordset.Fields("M_LOGO") = "AC" Or _
            '        adoReg.Recordset.Fields("M_LOGO") = "SV" Or _
            '        adoReg.Recordset.Fields("M_LOGO") = "LB" Or _
            '        adoReg.Recordset.Fields("M_LOGO") = "L4" Or _
            '        adoReg.Recordset.Fields("M_LOGO") = "SI" Or _
            '        adoReg.Recordset.Fields("M_LOGO") = "SY" Or _
            '        adoReg.Recordset.Fields("M_LOGO") = "CJ") Then

        
            '**** BONUS DE META
            'COD da loja do cx do ANO atual
            'Set w_ado_venda = de.cnc.Execute("SELECT TAB_VENDA.V_VR From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND TAB_FUNCIONARIO.F_DT_DEM IS NULL AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & Format(adoReg.Recordset.Fields("M_MES"), "00") & " AND Right(TAB_VENDA.V_DATA,4) = " & adoReg.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & adoReg.Recordset.Fields("M_F_COD")).Clone
            Set w_ado_venda = de.cnc.Execute("SELECT TAB_VENDA.V_VR FROM TAB_VENDA, TAB_FUNCIONARIO INNER JOIN Lojb010 ON TAB_FUNCIONARIO.F_Cod_L = Lojb010.COD_LOJ WHERE (((TAB_FUNCIONARIO.F_DT_DEM) Is Null) AND ((Right(Left([TAB_VENDA].[V_DATA],5),2))=" & Format(ADOREG.Recordset.Fields("M_MES"), "00") & ") AND ((Right([TAB_VENDA].[V_DATA],4))=" & ADOREG.Recordset.Fields("M_ANO") & ") AND ((TAB_FUNCIONARIO.F_Codigo)=" & ADOREG.Recordset.Fields("M_F_COD") & ") AND ((Lojb010.NUM)=Right([tab_venda].[v_f_loja],2)));").Clone
            If Not w_ado_venda.EOF Then
                vrVenda = w_ado_venda.Fields(0)
            Else
                vrVenda = 0
                'MsgBox "Não há lançamentos do logo " & adoReg.Recordset.Fields("M_LOGO") & " para o período: " & Format(adoReg.Recordset.Fields("M_MES"), "00") & " / " & adoReg.Recordset.Fields("M_ANO") & "! Ignorando...", vbCritical
            End If
    
            vrMeta = 999
            If ADOREG.Recordset.Fields("M_LOGO") <> "XX" And ADOREG.Recordset.Fields("M_LOGO") <> "99" Then
                vrMeta = de.cnc.Execute("SELECT TAB_META.MT_VALOR FROM TAB_META, Lojb010 INNER JOIN TAB_FUNCIONARIO ON Lojb010.COD_LOJ = TAB_FUNCIONARIO.F_Cod_L WHERE (((TAB_META.MT_MES)=" & ADOREG.Recordset.Fields("M_MES") & ") AND ((TAB_META.MT_ANO)=" & ADOREG.Recordset.Fields("M_ANO") & ") AND ((TAB_FUNCIONARIO.F_Codigo)=" & ADOREG.Recordset.Fields("M_F_COD") & ") AND ((Lojb010.NUM)=Right([tab_meta].[mt_f_loja],2)))").Fields(0)
            End If
            
            '****
            
            If ADOREG.Recordset.Fields("F_COD_CENTRAL") <> "" And IsNull(ADOREG.Recordset.Fields("M_DT_DEM")) Then

                
                adoComis.Filter = "F_4023717930 = " & ADOREG.Recordset.Fields("F_COD_CENTRAL")
                If Not adoComis.EOF Then
                    w_Piso = 0
                    w_Piso = de.cnc.Execute("SELECT F_VPISO_R FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
                    If w_Piso = 0 Then
                        w_Piso = de.cnc.Execute("SELECT F_VPISO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
                    End If
                    
                    w_Comis = 0
                    w_Premio = 0
                    w_Comis = CDbl(adoComis.Fields("COMTOTAL"))
                    w_Premio = CDbl(adoComis.Fields("F_1373503546"))
                    
                    de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & " AND (C_TP_CONTA = 20 OR C_TP_CONTA = 21 OR C_TP_CONTA = 23 OR C_TP_CONTA = 102 OR C_TP_CONTA = 101)")
                    
                'If adoReg.Recordset.Fields("M_MES") > "12" Then
                '        w_Fixo = de.cnc.Execute("SELECT IIf(IsNull(Lojb010.FixoVend),0,Lojb010.FixoVend) AS Fixo, TAB_FICHA_MENS.M_MES, TAB_FICHA_MENS.M_ANO FROM TAB_FICHA_MENS INNER JOIN Lojb010 ON TAB_FICHA_MENS.M_LOGO = Lojb010.COD_LOJ " & _
                '             " WHERE (((TAB_FICHA_MENS.M_F_COD)=" & adoReg.Recordset.Fields("M_F_COD") & ") " & _
                '             " AND ((TAB_FICHA_MENS.M_MES)=" & adoReg.Recordset.Fields("M_MES") & ") AND ((TAB_FICHA_MENS.M_ANO)=" & adoReg.Recordset.Fields("M_ANO") & "));").Fields(0)
                'Else
                    w_Fixo = 0
                'End If
                    
                  If (w_Comis + w_Premio + w_Fixo) <= w_Piso Then
                     'w_ultDiaMes = 30
                     w_ultDiaMes = Day(UltDiaMes(ADOREG.Recordset.Fields("m_mes"), ADOREG.Recordset.Fields("m_ano")))
                     If w_ultDiaMes < 30 Then w_ultDiaMes = 30
                 
                     'If adoReg.Recordset.Fields("m_dt_reg") = "" Or IsNull(adoReg.Recordset.Fields("m_dt_reg")) Then
                         w_Dt = CVDate(ADOREG.Recordset.Fields("m_dt_adm"))
                     'Else
                     '    w_Dt = CVDate(adoReg.Recordset.Fields("m_dt_reg"))
                     'End If
                     
                     
                     If IsDate(ADOREG.Recordset.Fields("M_DT_DEM")) Then
                         w_dtUlt = CVDate(ADOREG.Recordset.Fields("M_DT_DEM"))
                     Else
                         w_dtUlt = CVDate(UltDiaMes(ADOREG.Recordset.Fields("m_mes"), ADOREG.Recordset.Fields("m_ano")))
                         'If Day(w_dtUlt) = 31 Then w_dtUlt = w_dtUlt - 1
                         If Day(w_dtUlt) = 28 Then w_dtUlt = w_dtUlt + 2
                         If Day(w_dtUlt) = 29 Then w_dtUlt = w_dtUlt + 1
                     End If
                     
                     'If Month(w_Dt) < Month(w_dtUlt) Then w_Dt = CVDate("01/" & Month(w_dtUlt) & "/" & Year(w_dtUlt))
                     
                     w_DtDiff = DateDiff("d", w_Dt, w_dtUlt) + 1
                     
                     w_PisoOriginal = w_Piso
                     w_FixoOriginal = w_Fixo
                     'MsgBox "Diff: " & w_DtDiff & " - Ini: " & w_Dt & " - Final: " & w_dtUlt
                     If (w_DtDiff < w_ultDiaMes) And (w_DtDiff < 30) Then
                         If w_ultDiaMes < 30 Then
                            If w_Fixo > 0 Then w_Fixo = ((w_Fixo / w_ultDiaMes) * w_DtDiff)
                            w_Piso = ((w_Piso / w_ultDiaMes) * w_DtDiff)
                         Else
                            If w_Fixo > 0 Then w_Fixo = ((w_Fixo / 30) * w_DtDiff)
                            w_Piso = ((w_Piso / 30) * w_DtDiff)
                            w_ultDiaMes = 30
                         End If
                         
                         If w_FixoOriginal > 0 Then
                            If (w_Comis + w_Premio + w_Fixo) <= w_Piso Then
                                de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 23, "+", Format(w_Piso, "0.00"), "#NÃO ATINGIU O PISO# Fixo: " & Format(w_FixoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_FixoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Fixo, "0.00") & " + Comissão: " & Format(w_Comis, "0.00") & " + Prêmio: " & Format(w_Premio, "0.00") & " = " & Format(w_Fixo + w_Comis + w_Premio, "0.00") & ". Piso: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                            Else
                                de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 20, "+", Format(w_Comis, "0.00"), "#SALÁRIO MAIOR QUE PISO PROPORCIONAL# Comissão: " & Format(w_Comis, "0.00") & ". Piso: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                                de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 102, "+", Format(w_Fixo, "0.00"), "#SALÁRIO MAIOR QUE PISO PROPORCIONAL# Fixo: " & Format(w_FixoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_FixoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Fixo, "0.00") & ". Piso: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                                de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 21, "+", Format(w_Premio, "0.00"), "PRÊMIO [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                            End If
                         Else
                            If (w_Comis + w_Premio) <= w_Piso Then
                                de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 23, "+", Format(w_Piso, "0.00"), "#NÃO ATINGIU O PISO# Comissão: " & Format(w_Comis, "0.00") & " + Prêmio: " & Format(w_Premio, "0.00") & " = " & Format(w_Comis + w_Premio, "0.00") & ". Piso: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                            Else
                                de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 20, "+", Format(w_Comis, "0.00"), "#COMISSÃO MAIOR QUE PISO PROPORCIONAL# Comissão: " & Format(w_Comis, "0.00") & " + Prêmio: " & Format(w_Premio, "0.00") & " = " & Format(w_Comis + w_Premio, "0.00") & ". Piso: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                                de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 21, "+", Format(w_Premio, "0.00"), "PRÊMIO [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                            End If
                         End If
        
                     Else
                        If w_FixoOriginal > 0 Then
                            de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 23, "+", Format(w_Piso, "0.00"), "#NÃO ATINGIU O PISO# Fixo: " & Format(w_Fixo, "0.00") & " + Comissão: " & Format(w_Comis, "0.00") & " + Prêmio: " & Format(w_Premio, "0.00") & " = " & Format(w_Fixo + w_Comis + w_Premio, "0.00") & ". Piso: " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                        Else
                            de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 23, "+", Format(w_Piso, "0.00"), "#NÃO ATINGIU O PISO# Comissão: " & Format(w_Comis, "0.00") & " + Prêmio: " & Format(w_Premio, "0.00") & " = " & Format(w_Comis + w_Premio, "0.00") & ". Piso: " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                        End If
                     End If
                     
                 Else
                    If w_Fixo > 0 Then
                       de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 20, "+", Format(w_Comis, "0.00"), "COMISSÃO [GERADO AUTOMATICAMENTE] Comissão: " & Format(w_Comis, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                       de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 102, "+", Format(w_Fixo, "0.00"), "FIXO [GERADO AUTOMATICAMENTE] Fixo: " & Format(w_Fixo, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                       de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 21, "+", Format(w_Premio, "0.00"), "PRÊMIO [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                    Else
                       de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 20, "+", Format(w_Comis, "0.00"), "COMISSÃO [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                       de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 21, "+", Format(w_Premio, "0.00"), "PRÊMIO [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                    End If
                 End If
            End If
           ElseIf IsNull(ADOREG.Recordset.Fields("M_DT_DEM")) Then
           
            'Dim w_Dt, w_dtUlt As Date
            'Dim w_DtDiff, w_ultDiaMes As Integer
            'Dim w_Piso, w_Comis, w_Premio, w_PisoOriginal As Double
                
                If Not adoComis.EOF Then
                    w_Piso = 0
                    w_Piso = de.cnc.Execute("SELECT F_VPISO_R FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
                    If w_Piso = 0 Then
                        w_Piso = de.cnc.Execute("SELECT F_VPISO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
                    End If
                    
                    'If adoReg.Recordset.Fields("M_MES") = "1" Or adoReg.Recordset.Fields("M_MES") = "2" Or adoReg.Recordset.Fields("M_MES") = "3" Then
                    '    w_Fixo = de.cnc.Execute("SELECT IIf(IsNull(Lojb010.FixoVend),0,Lojb010.FixoVend) AS Fixo, TAB_FICHA_MENS.M_MES, TAB_FICHA_MENS.M_ANO FROM TAB_FICHA_MENS INNER JOIN Lojb010 ON TAB_FICHA_MENS.M_LOGO = Lojb010.COD_LOJ " & _
                    '         " WHERE (((TAB_FICHA_MENS.M_F_COD)=" & adoReg.Recordset.Fields("M_F_COD") & ") " & _
                    '         " AND ((TAB_FICHA_MENS.M_MES)=" & adoReg.Recordset.Fields("M_MES") & ") AND ((TAB_FICHA_MENS.M_ANO)=" & adoReg.Recordset.Fields("M_ANO") & "));").Fields(0)
                    'Else
                        w_Fixo = 0
                    'End If
                    
                    de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & " AND (C_TP_CONTA = 20 OR C_TP_CONTA = 21 OR C_TP_CONTA = 23  OR C_TP_CONTA = 102)")
                    
                     'w_ultDiaMes = 30
                     w_ultDiaMes = Day(UltDiaMes(ADOREG.Recordset.Fields("m_mes"), ADOREG.Recordset.Fields("m_ano")))
                     If w_ultDiaMes < 30 Then w_ultDiaMes = 30
                 
                     'If adoReg.Recordset.Fields("m_dt_reg") = "" Or IsNull(adoReg.Recordset.Fields("m_dt_reg")) Then
                         w_Dt = CVDate(ADOREG.Recordset.Fields("m_dt_adm"))
                    ' Else
                     '    w_Dt = CVDate(adoReg.Recordset.Fields("m_dt_reg"))
                    ' End If
                     
                     
                     If IsDate(ADOREG.Recordset.Fields("M_DT_DEM")) Then
                         w_dtUlt = CVDate(ADOREG.Recordset.Fields("M_DT_DEM"))
                     Else
                         w_dtUlt = CVDate(UltDiaMes(ADOREG.Recordset.Fields("m_mes"), ADOREG.Recordset.Fields("m_ano")))
                         'If Day(w_dtUlt) = 31 Then w_dtUlt = w_dtUlt - 1
                         If Day(w_dtUlt) = 28 Then w_dtUlt = w_dtUlt + 2
                         If Day(w_dtUlt) = 29 Then w_dtUlt = w_dtUlt + 1
                     End If
                     
                     'If Month(w_Dt) < Month(w_dtUlt) Then w_Dt = CVDate("01/" & Month(w_dtUlt) & "/" & Year(w_dtUlt))
                     
                     w_DtDiff = DateDiff("d", w_Dt, w_dtUlt) + 1
                     
                     w_PisoOriginal = w_Piso
                     w_FixoOriginal = w_Fixo
                     If (w_DtDiff < w_ultDiaMes) And (w_DtDiff < 30) Then
                         If w_ultDiaMes < 30 Then
                            If w_Fixo > 0 Then w_Fixo = ((w_Fixo / w_ultDiaMes) * w_DtDiff)
                            w_Piso = ((w_Piso / w_ultDiaMes) * w_DtDiff)
                         Else
                            If w_Fixo > 0 Then w_Fixo = ((w_Fixo / 30) * w_DtDiff)
                            w_Piso = ((w_Piso / 30) * w_DtDiff)
                            w_ultDiaMes = 30
                         End If
                        If w_FixoOriginal > 0 Then
                           de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 23, "+", Format(w_Piso, "0.00"), "#NÃO ATINGIU O PISO PROPORCIONAL# Fixo: " & Format(w_Fixo, "0.00") & " + Comissão: " & Format(0, "0.00") & " + Prêmio: " & Format(0, "0.00") & " = " & Format(0 + 0, "0.00") & ". Piso: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                        Else
                           de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 23, "+", Format(w_Piso, "0.00"), "#NÃO ATINGIU O PISO PROPORCIONAL# Comissão: " & Format(0, "0.00") & " + Prêmio: " & Format(0, "0.00") & " = " & Format(0 + 0, "0.00") & ". Piso: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                        End If
                     
                     Else
                        If w_FixoOriginal > 0 Then
                            de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 23, "+", Format(w_Piso, "0.00"), "#NÃO ATINGIU O PISO# Fixo: " & Format(w_Fixo, "0.00") & " + Comissão: " & Format(0, "0.00") & " + Prêmio: " & Format(0, "0.00") & " = " & Format(0 + 0, "0.00") & ". Piso: " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                        Else
                            de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 23, "+", Format(w_Piso, "0.00"), "#NÃO ATINGIU O PISO# Comissão: " & Format(0, "0.00") & " + Prêmio: " & Format(0, "0.00") & " = " & Format(0 + 0, "0.00") & ". Piso: " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                        End If
                     End If
            End If
            
           End If
           
             vrSalario = w_Comis + w_Premio + w_Fixo
             vrMetaBonus = 0
             wFixoMeta = 0
             wPercMeta = 0
                
            If IsNull(ADOREG.Recordset.Fields("M_DT_DEM")) Then
             
                If vrVenda >= (vrMeta) Then
                
                    If ADOREG.Recordset.Fields("m_mes") = 1 Then

                        If w_Piso > vrSalario Then
                           vrSalario = w_Piso
                        End If
                    End If
                
                    wFixoMeta = 150
                    If (((vrVenda / vrMeta) - 1) * 100) >= 5 Then
                        vrMetaBonus = vrSalario * 0.05
                        wPercMeta = 5
                    Else
                        vrMetaBonus = vrSalario * ((vrVenda / vrMeta) - 1)
                        wPercMeta = ((vrVenda / vrMeta) - 1) * 100
                    End If

                    de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & " AND (C_TP_CONTA = 101)")
                    'Bonus de % do Salário
                    de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 101, "+", Format(vrMetaBonus + wFixoMeta, "0.00"), "#BONUS DE META# Fixo(" & Format(w_Fixo, "0.00") & ") + Comissão(" & Format(w_Comis, "0.00") & ") + Prêmio(" & Format(w_Premio, "0.00") & ") = " & Format(vrSalario, "0.00") & " * " & Format(wPercMeta, "0.0") & "% = " & Format(vrMetaBonus, "0.00") & " + Fixo Meta(" & Format(wFixoMeta, "0.00") & ") = " & Format(vrMetaBonus + wFixoMeta, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                    vrSalario = vrSalario + vrMetaBonus + wFixoMeta
                Else
                   descMeta = ""
                End If
                

             
            End If
           
           End If
           
           'End If

           ADOREG.Recordset.MoveNext
           
            
        Loop
        
        ADOREG.Recordset.MoveFirst
        Lancamentos
        
        adoComis.Clone
    Else
        MsgBox "Comissão cancelada, exibindo apenas o relatório...", vbInformation, "Exibindo relatório"
    End If

    de.cmdSqlDeleteComissTMP
    de.cmdComissDump

    Set adoComisDump = de.rscmdComissDump.Clone
    
    If Not adoComisDump.EOF Then
        adoComisDump.MoveFirst
    End If
    
    Do While Not adoComisDump.EOF
    de.cmdSqlAddComissTMP IIf(IsNull(adoComisDump.Fields("cod_filial")), "", adoComisDump.Fields("cod_filial")), _
        IIf(IsNull(adoComisDump.Fields("F_4023717930")), "", adoComisDump.Fields("F_4023717930")), _
        IIf(IsNull(adoComisDump.Fields("funcionario_GERADOR_NOME")), "", adoComisDump.Fields("funcionario_GERADOR_NOME")), _
        IIf(IsNull(adoComisDump.Fields("filial_filial_filial")), "", adoComisDump.Fields("filial_filial_filial")), _
        IIf(IsNull(adoComisDump.Fields("funcionario_GERADOR_GERADOR")), "", adoComisDump.Fields("funcionario_GERADOR_GERADOR")), _
        IIf(IsNull(adoComisDump.Fields("F_2466906013")), 0, adoComisDump.Fields("F_2466906013")), _
        IIf(IsNull(adoComisDump.Fields("F_3050360625")), 0, adoComisDump.Fields("F_3050360625")), _
        IIf(IsNull(adoComisDump.Fields("F_1373503546")), 0, adoComisDump.Fields("F_1373503546")), _
        IIf(IsNull(adoComisDump.Fields("vendTotal")), 0, adoComisDump.Fields("vendTotal")), _
        IIf(IsNull(adoComisDump.Fields("comVista")), 0, adoComisDump.Fields("comVista")), _
        IIf(IsNull(adoComisDump.Fields("comPrazo")), 0, adoComisDump.Fields("comPrazo")), _
        IIf(IsNull(adoComisDump.Fields("comTotal")), 0, adoComisDump.Fields("comTotal")), _
        IIf(IsNull(adoComisDump.Fields("totGeral")), 0, adoComisDump.Fields("totGeral"))
        'de.cmdSqlAddComissTMP IIf(IsNull(adoComisDump.Fields(14)), "", adoComisDump.Fields(14)), IIf(IsNull(adoComisDump.Fields(1)), "", adoComisDump.Fields(1)), IIf(IsNull(adoComisDump.Fields(2)), "", adoComisDump.Fields(2)), IIf(IsNull(adoComisDump.Fields(3)), "", adoComisDump.Fields(3)), IIf(IsNull(adoComisDump.Fields(4)), "", adoComisDump.Fields(4)), IIf(IsNull(adoComisDump.Fields(5)), 0, adoComisDump.Fields(5)), IIf(IsNull(adoComisDump.Fields(6)), 0, adoComisDump.Fields(6)), IIf(IsNull(adoComisDump.Fields(7)), 0, adoComisDump.Fields(7)), IIf(IsNull(adoComisDump.Fields(8)), 0, adoComisDump.Fields(8)), IIf(IsNull(adoComisDump.Fields(9)), 0, adoComisDump.Fields(9)), IIf(IsNull(adoComisDump.Fields(10)), 0, adoComisDump.Fields(10)), IIf(IsNull(adoComisDump.Fields(11)), 0, adoComisDump.Fields(11)), IIf(IsNull(adoComisDump.Fields(12)), 0, adoComisDump.Fields(12))
        'de.cmdSqlAddComissTMP IIf(IsNull(adoComisDump.Fields(14)), "", adoComisDump.Fields(0)), IIf(IsNull(adoComisDump.Fields(1)), "", adoComisDump.Fields(1)), IIf(IsNull(adoComisDump.Fields(2)), "", adoComisDump.Fields(2)), IIf(IsNull(adoComisDump.Fields(3)), "", adoComisDump.Fields(3)), IIf(IsNull(adoComisDump.Fields(4)), "", adoComisDump.Fields(4)), IIf(IsNull(adoComisDump.Fields(5)), 0, adoComisDump.Fields(5)), IIf(IsNull(adoComisDump.Fields(6)), 0, adoComisDump.Fields(6)), IIf(IsNull(adoComisDump.Fields(7)), 0, adoComisDump.Fields(7)), IIf(IsNull(adoComisDump.Fields(8)), 0, adoComisDump.Fields(8)), IIf(IsNull(adoComisDump.Fields(9)), 0, adoComisDump.Fields(9)), IIf(IsNull(adoComisDump.Fields(10)), 0, adoComisDump.Fields(10)), IIf(IsNull(adoComisDump.Fields(11)), 0, adoComisDump.Fields(11)), IIf(IsNull(adoComisDump.Fields(12)), 0, adoComisDump.Fields(12))
        adoComisDump.MoveNext
    Loop

    sqlComissTMP = "SELECT TAB_COMISSAO_TEMP.filial_filial_NOME, TAB_COMISSAO_TEMP.F_4023717930, " & _
             "TAB_COMISSAO_TEMP.funcionario_GERADOR_NOME, TAB_COMISSAO_TEMP.filial_filial_filial,  " & _
             "TAB_COMISSAO_TEMP.funcionario_GERADOR_GERADOR, TAB_COMISSAO_TEMP.F_2466906013, TAB_COMISSAO_TEMP.F_3050360625,  " & _
             "TAB_COMISSAO_TEMP.F_1373503546, TAB_COMISSAO_TEMP.vendTotal, TAB_COMISSAO_TEMP.comVista, TAB_COMISSAO_TEMP.comPrazo,  " & _
             "TAB_COMISSAO_TEMP.comTotal, TAB_COMISSAO_TEMP.totGeral, (SELECT Max(TAB_FICHA_MENS.M_DT_DEM)  " & _
             "FROM TAB_FUNCIONARIO INNER JOIN TAB_FICHA_MENS ON TAB_FUNCIONARIO.F_Codigo = TAB_FICHA_MENS.M_F_COD " & _
             "WHERE (((TAB_FICHA_MENS.M_ANO)=" & TXT_ANO & ") AND ((TAB_FICHA_MENS.M_MES)=" & TXT_MES & ") AND " & _
             "((TAB_FUNCIONARIO.F_COD_CENTRAL)=[TAB_COMISSAO_TEMP].[F_4023717930]) AND ((TAB_FICHA_MENS.M_TIPO)='V'))) AS DT_DEM, " & _
             "(SELECT Sum((C_VALOR)) FROM TAB_FUNCIONARIO INNER JOIN (TAB_FICHA_MENS INNER JOIN TAB_DESC_CALC ON TAB_FICHA_MENS.M_NFICHA " & _
             "= TAB_DESC_CALC.C_N_FICHA) ON TAB_FUNCIONARIO.F_Codigo = TAB_FICHA_MENS.M_F_COD " & _
             "WHERE (((TAB_DESC_CALC.C_TP_CONTA)=20 Or (TAB_DESC_CALC.C_TP_CONTA)=21 Or (TAB_DESC_CALC.C_TP_CONTA)=23) AND ((TAB_FICHA_MENS.M_TIPO)='V')) " & _
             "GROUP BY TAB_FICHA_MENS.M_ANO, TAB_FICHA_MENS.M_MES, TAB_FUNCIONARIO.F_COD_CENTRAL, TAB_FUNCIONARIO.F_COD_CENTRAL " & _
             "HAVING (((TAB_FICHA_MENS.M_ANO)=" & TXT_ANO & ") AND ((TAB_FICHA_MENS.M_MES)=" & TXT_MES & ") AND " & _
             "((TAB_FUNCIONARIO.F_COD_CENTRAL)=[TAB_COMISSAO_TEMP].[F_4023717930]))) AS Salario  " & _
             "FROM TAB_COMISSAO_TEMP"
    
    de.rscmdComissTMP.Open sqlComissTMP, , adOpenStatic, adLockOptimistic
    Set adoComisTMP = de.rscmdComissTMP.Clone

    If Not adoComisTMP.EOF Then
        adoComisTMP.MoveFirst
        Do While Not adoComisTMP.EOF
        de.cmdSqlUpdComissTMP IIf(IsNull(adoComisTMP.Fields("DT_DEM")), "", adoComisTMP.Fields("DT_DEM")), _
            IIf(IsNull(adoComisTMP.Fields("Salario")), 0, adoComisTMP.Fields("Salario")), _
            IIf(IsNull(adoComisTMP.Fields("filial_filial_NOME")), "", adoComisTMP.Fields("filial_filial_NOME")), _
            IIf(IsNull(adoComisTMP.Fields("F_4023717930")), "", adoComisTMP.Fields("F_4023717930")), _
            IIf(IsNull(adoComisTMP.Fields("funcionario_GERADOR_NOME")), "", adoComisTMP.Fields("funcionario_GERADOR_NOME")), _
            IIf(IsNull(adoComisTMP.Fields("filial_filial_filial")), "", adoComisTMP.Fields("filial_filial_filial")), _
            IIf(IsNull(adoComisTMP.Fields("funcionario_GERADOR_GERADOR")), "", adoComisTMP.Fields("funcionario_GERADOR_GERADOR")), _
            IIf(IsNull(adoComisTMP.Fields("F_2466906013")), 0, adoComisTMP.Fields("F_2466906013")), _
            IIf(IsNull(adoComisTMP.Fields("F_3050360625")), 0, adoComisTMP.Fields("F_3050360625")), _
            IIf(IsNull(adoComisTMP.Fields("F_1373503546")), 0, adoComisTMP.Fields("F_1373503546")), _
            IIf(IsNull(adoComisTMP.Fields("vendTotal")), 0, adoComisTMP.Fields("vendTotal")), _
            IIf(IsNull(adoComisTMP.Fields("comVista")), 0, adoComisTMP.Fields("comVista")), _
            IIf(IsNull(adoComisTMP.Fields("comPrazo")), 0, adoComisTMP.Fields("comPrazo")), _
            IIf(IsNull(adoComisTMP.Fields("comTotal")), 0, adoComisTMP.Fields("comTotal")), _
            IIf(IsNull(adoComisTMP.Fields("totGeral")), 0, adoComisTMP.Fields("totGeral"))
            'de.cmdSqlUpdComissTMP IIf(IsNull(adoComisTMP.Fields(13)), "", adoComisTMP.Fields(13)), IIf(IsNull(adoComisTMP.Fields(14)), 0, adoComisTMP.Fields(14)), IIf(IsNull(adoComisTMP.Fields(0)), "", adoComisTMP.Fields(0)), IIf(IsNull(adoComisTMP.Fields(1)), "", adoComisTMP.Fields(1)), IIf(IsNull(adoComisTMP.Fields(2)), "", adoComisTMP.Fields(2)), IIf(IsNull(adoComisTMP.Fields(3)), "", adoComisTMP.Fields(3)), IIf(IsNull(adoComisTMP.Fields(4)), "", adoComisTMP.Fields(4)), IIf(IsNull(adoComisTMP.Fields(5)), 0, adoComisTMP.Fields(5)), IIf(IsNull(adoComisTMP.Fields(6)), 0, adoComisTMP.Fields(6)), IIf(IsNull(adoComisTMP.Fields(7)), 0, adoComisTMP.Fields(7)), IIf(IsNull(adoComisTMP.Fields(8)), 0, adoComisTMP.Fields(8)), IIf(IsNull(adoComisTMP.Fields(9)), 0, adoComisTMP.Fields(9)), IIf(IsNull(adoComisTMP.Fields(10)), 0, adoComisTMP.Fields(10)), IIf(IsNull(adoComisTMP.Fields(11)), 0, adoComisTMP.Fields(11)), IIf(IsNull(adoComisTMP.Fields(12)), 0, adoComisTMP.Fields(12))
            adoComisTMP.MoveNext
        Loop
    End If

    'If adoReg.Recordset.Fields("M_MES") > "12" Then
    '    If de.rscmdRptComissTMPFixo_Grouping.State = 1 Then de.rscmdRptComissTMPFixo_Grouping.Close
    '    de.cmdRptComissTMPFixo_Grouping
    '    rptComissMwtsFixo.Show
    'Else
        If de.rscmdRptComissTMP_Grouping.State = 1 Then de.rscmdRptComissTMP_Grouping.Close
        de.cmdRptComissTMP_Grouping
        rptComissMwts.Show
    'End If
    
    'adoComis.Close
    'de.rscmdComissGerar.Close
    
    
End Sub

Private Sub cmdComissVendedor_Click()
Dim dtIni, dtFim As Date
Dim adoComis As ADODB.Recordset

    frm_ESCOLHA_DATA.Show 1
    
    dtIni = CDate(frm_ESCOLHA_DATA.TXT_DT_INICIAL)
    dtFim = CDate(frm_ESCOLHA_DATA.TXT_DT_FINAL)
     
    'On Error Resume Next
    'de.cmdDROPtmpComis1
    'de.cmdDROPtmpComis2
    
    'de.cmdCREATEtmpComis1
    'de.cmdCREATEtmpComis2
    
    de.cmdDELETEtmpComis1
    de.cmdDELETEtmpComis2
    
    de.cmdAddtmpComis1 dtIni, dtFim, dtIni, dtFim, dtIni, dtFim, dtIni, dtFim, dtIni, dtFim
    de.cmdAddtmpComis2 dtIni, dtFim, dtIni, dtFim
   
     de.cmdComissGerar
     Set adoComis = de.rscmdComissGerar.Clone
         
     If ADOREG.Recordset.Fields("M_TIPO") = "V" And ADOREG.Recordset.Fields("F_COD_CENTRAL") <> "" And IsNull(ADOREG.Recordset.Fields("M_DT_DEM")) Then
         Dim w_Dt, w_dtUlt As Date
         Dim w_DtDiff, w_ultDiaMes As Integer
         Dim w_Piso, w_Comis, w_Premio, w_PisoOriginal As Double
         
         adoComis.Filter = "F_4023717930 = " & ADOREG.Recordset.Fields("F_COD_CENTRAL")
         If Not adoComis.EOF Then
             w_Piso = 0
             w_Piso = de.cnc.Execute("SELECT F_VPISO_R FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
             If w_Piso = 0 Then
                 w_Piso = de.cnc.Execute("SELECT F_VPISO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
             End If
             
             w_Comis = CDbl(adoComis.Fields("COMTOTAL"))
             w_Premio = CDbl(adoComis.Fields("F_1373503546"))
             
             de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & " AND (C_TP_CONTA = 20 OR C_TP_CONTA = 21 OR C_TP_CONTA = 23)")
             
             If (adoComis.Fields("COMTOTAL") + adoComis.Fields("F_1373503546")) <= w_Piso Then
                 'w_ultDiaMes = 30
                 w_ultDiaMes = Day(UltDiaMes(ADOREG.Recordset.Fields("m_mes"), ADOREG.Recordset.Fields("m_ano")))
                 If w_ultDiaMes < 30 Then w_ultDiaMes = 30
             
                 If ADOREG.Recordset.Fields("m_dt_reg") = "" Or IsNull(ADOREG.Recordset.Fields("m_dt_reg")) Then
                     w_Dt = CVDate(ADOREG.Recordset.Fields("m_dt_adm"))
                 Else
                     w_Dt = CVDate(ADOREG.Recordset.Fields("m_dt_reg"))
                 End If
                 
                 
                 If IsDate(ADOREG.Recordset.Fields("M_DT_DEM")) Then
                     w_dtUlt = CVDate(ADOREG.Recordset.Fields("M_DT_DEM"))
                 Else
                     w_dtUlt = CVDate(UltDiaMes(ADOREG.Recordset.Fields("m_mes"), ADOREG.Recordset.Fields("m_ano")))
                     'If Day(w_dtUlt) = 31 Then w_dtUlt = w_dtUlt - 1
                     If Day(w_dtUlt) = 28 Then w_dtUlt = w_dtUlt + 2
                     If Day(w_dtUlt) = 29 Then w_dtUlt = w_dtUlt + 1
                 End If
                 
                 If Month(w_Dt) < Month(w_dtUlt) Then w_Dt = CVDate("01/" & Month(w_dtUlt) & "/" & Year(w_dtUlt))
                 
                 w_DtDiff = DateDiff("d", w_Dt, w_dtUlt) + 1
                 
                 w_PisoOriginal = w_Piso
                 'MsgBox "Diff: " & w_DtDiff & " - Ini: " & w_Dt & " - Final: " & w_dtUlt
                 If (w_DtDiff < w_ultDiaMes) And (w_DtDiff < 30) Then
                     If w_ultDiaMes < 30 Then
                         w_Piso = ((w_Piso / w_ultDiaMes) * w_DtDiff)
                     Else
                         w_Piso = ((w_Piso / 30) * w_DtDiff)
                         w_ultDiaMes = 30
                     End If
                     
                     If (adoComis.Fields("COMTOTAL") + adoComis.Fields("F_1373503546")) <= w_Piso Then
                         de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 23, "+", Format(w_Piso, "0.00"), "#NÃO ATINGIU O PISO# Comissão: " & Format(w_Comis, "0.00") & " + Prêmio: " & Format(w_Premio, "0.00") & " = " & Format(w_Comis + w_Premio, "0.00") & ". Piso: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                     Else
                         de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 20, "+", Format(w_Comis, "0.00"), "#COMISSÃO MAIOR QUE PISO PROPORCIONAL# Comissão: " & Format(w_Comis, "0.00") & " + Prêmio: " & Format(w_Premio, "0.00") & " = " & Format(w_Comis + w_Premio, "0.00") & ". Piso: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                         de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 21, "+", Format(w_Premio, "0.00"), "PRÊMIO [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                     End If
    
                 Else
                     de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 23, "+", Format(w_Piso, "0.00"), "#NÃO ATINGIU O PISO# Comissão: " & Format(w_Comis, "0.00") & " + Prêmio: " & Format(w_Premio, "0.00") & " = " & Format(w_Comis + w_Premio, "0.00") & ". Piso: " & Format(w_Piso, "0.00"), ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                 End If
                 
             Else
                 de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 20, "+", Format(w_Comis, "0.00"), "COMISSÃO [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
                 de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 21, "+", Format(w_Premio, "0.00"), "PRÊMIO [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
             End If
         End If
     End If
     
    adoComis.Close
    de.rscmdComissGerar.Close
    Lancamentos

    

End Sub


Private Sub cmdComissaoVendedor_Click()
    'If adoReg.Recordset.Fields("M_MES") = "1" Or adoReg.Recordset.Fields("M_MES") = "2" Or adoReg.Recordset.Fields("M_MES") = "3" Then
    '    If de.rscmdRptComissTMPFixo_Grouping.State = 1 Then de.rscmdRptComissTMPFixo_Grouping.Close
    '    de.cmdRptComissTMPFixo_Grouping
    '    rptComissMwtsFixo.Show
    'Else
        If de.rscmdRptComissTMP_Grouping.State = 1 Then de.rscmdRptComissTMP_Grouping.Close
        de.cmdRptComissTMP_Grouping
        rptComissMwts.Show
    'End If
End Sub

Private Sub cmdComixCxEXT_Click()
Dim dtIni, dtFim As Date
Dim vrMeta, vrMetaBonus, descMeta, vrPisoMeta
   
    If de.rscmdComissGerar.State = 1 Then de.rscmdComissGerar.Close
    de.cmdComissGerar
    Set adoComis = de.rscmdComissGerar.Clone
        
    cbMostrar.text = "CX EXTRA"
    cmdMostrar_Click
        
    ADOREG.Recordset.MoveFirst
    Do While Not ADOREG.Recordset.EOF
        If (ADOREG.Recordset.Fields("M_TIPO") = "X" Or ADOREG.Recordset.Fields("M_TIPO") = "2") And IsNull(ADOREG.Recordset.Fields("M_DT_DEM")) Then
            Dim w_Dt, w_dtUlt As Date
            Dim w_DtDiff, w_ultDiaMes As Integer
            Dim w_Piso, w_Comis, w_Premio, w_PisoOriginal As Double
            
            w_Piso = 0
            w_Piso = de.cnc.Execute("SELECT F_VR_FIXO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
            If IsNull(w_Piso) Or (w_Piso = 0) Then
            
                w_Piso = de.cnc.Execute("SELECT F_VPISO_R FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
                If w_Piso = 0 Then
                    w_Piso = de.cnc.Execute("SELECT F_VPISO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
                End If
            End If
                
            de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & " AND (C_TP_CONTA = 22)")
            
             'w_ultDiaMes = 30
             w_ultDiaMes = Day(UltDiaMes(ADOREG.Recordset.Fields("m_mes"), ADOREG.Recordset.Fields("m_ano")))
             If w_ultDiaMes < 30 Then w_ultDiaMes = 30
         
             If ADOREG.Recordset.Fields("m_dt_reg") = "" Or IsNull(ADOREG.Recordset.Fields("m_dt_reg")) Then
                 w_Dt = CVDate(ADOREG.Recordset.Fields("m_dt_adm"))
             Else
                 w_Dt = CVDate(ADOREG.Recordset.Fields("m_dt_reg"))
             End If
             
             
             If IsDate(ADOREG.Recordset.Fields("M_DT_DEM")) Then
                 w_dtUlt = CVDate(ADOREG.Recordset.Fields("M_DT_DEM"))
             Else
                 w_dtUlt = CVDate(UltDiaMes(ADOREG.Recordset.Fields("m_mes"), ADOREG.Recordset.Fields("m_ano")))
                 'If Day(w_dtUlt) = 31 Then w_dtUlt = w_dtUlt - 1
                 If Day(w_dtUlt) = 28 Then w_dtUlt = w_dtUlt + 2
                 If Day(w_dtUlt) = 29 Then w_dtUlt = w_dtUlt + 1
             End If
             
             If Month(w_Dt) < Month(w_dtUlt) Then w_Dt = CVDate("01/" & Month(w_dtUlt) & "/" & Year(w_dtUlt))
             
             w_DtDiff = DateDiff("d", w_Dt, w_dtUlt) + 1
             
             w_PisoOriginal = w_Piso
             'MsgBox "Diff: " & w_DtDiff & " - Ini: " & w_Dt & " - Final: " & w_dtUlt
             
             
             If (w_DtDiff < w_ultDiaMes) And (w_DtDiff < 30) Then
                 If w_ultDiaMes < 30 Then
                     w_Piso = ((w_Piso / w_ultDiaMes) * w_DtDiff)
                 Else
                     w_Piso = ((w_Piso / 30) * w_DtDiff)
                     w_ultDiaMes = 30
                 End If
             End If
             
             vrPisoMeta = w_Piso
             
             If (ADOREG.Recordset.Fields("M_TIPO") = "2") Then
             
                '**** BONUS DE META
                'COD da loja do cx do ANO atual
                'Set w_ado_venda = de.cnc.Execute("SELECT TAB_VENDA.V_VR From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND TAB_FUNCIONARIO.F_DT_DEM IS NULL AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & Format(adoReg.Recordset.Fields("M_MES"), "00") & " AND Right(TAB_VENDA.V_DATA,4) = " & adoReg.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & adoReg.Recordset.Fields("M_F_COD")).Clone
                Set w_ado_venda = de.cnc.Execute("SELECT TAB_VENDA.V_VR FROM TAB_VENDA, TAB_FUNCIONARIO INNER JOIN Lojb010 ON TAB_FUNCIONARIO.F_Cod_L = Lojb010.COD_LOJ WHERE (((TAB_FUNCIONARIO.F_DT_DEM) Is Null) AND ((Right(Left([TAB_VENDA].[V_DATA],5),2))=" & Format(ADOREG.Recordset.Fields("M_MES"), "00") & ") AND ((Right([TAB_VENDA].[V_DATA],4))=" & ADOREG.Recordset.Fields("M_ANO") & ") AND ((TAB_FUNCIONARIO.F_Codigo)=" & ADOREG.Recordset.Fields("M_F_COD") & ") AND ((Lojb010.NUM)=Right([tab_venda].[v_f_loja],2)));").Clone
                If Not w_ado_venda.EOF Then
                    vrVenda = w_ado_venda.Fields(0)
                Else
                    vrVenda = 0
                    'MsgBox "Não há lançamentos do logo " & adoReg.Recordset.Fields("M_LOGO") & " para o período: " & Format(adoReg.Recordset.Fields("M_MES"), "00") & " / " & adoReg.Recordset.Fields("M_ANO") & "! Ignorando...", vbCritical
                End If
        
                vrMeta = 999
                If ADOREG.Recordset.Fields("M_LOGO") <> "XX" Then
                    vrMeta = de.cnc.Execute("SELECT TAB_META.MT_VALOR FROM TAB_META, Lojb010 INNER JOIN TAB_FUNCIONARIO ON Lojb010.COD_LOJ = TAB_FUNCIONARIO.F_Cod_L WHERE (((TAB_META.MT_MES)=" & ADOREG.Recordset.Fields("M_MES") & ") AND ((TAB_META.MT_ANO)=" & ADOREG.Recordset.Fields("M_ANO") & ") AND ((TAB_FUNCIONARIO.F_Codigo)=" & ADOREG.Recordset.Fields("M_F_COD") & ") AND ((Lojb010.NUM)=Right([tab_meta].[mt_f_loja],2)))").Fields(0)
                End If
                
                 vrMetaBonus = 0
                
                If ADOREG.Recordset.Fields("m_mes") <= 12 Then
                    wPercMeta = 0
                    If vrVenda >= (vrMeta * 1.1) Then
                        vrMetaBonus = vrSalario * 0.1
                        wPercMeta = 10
                    ElseIf vrVenda >= (vrMeta * 1.08) Then
                        vrMetaBonus = vrSalario * 0.08
                        wPercMeta = 8
                    ElseIf vrVenda >= (vrMeta * 1.05) Then
                        vrMetaBonus = vrSalario * 0.05
                        wPercMeta = 5
                    End If
                    
                    If vrVenda >= (vrMeta * 1.05) Then
                        descMeta = " + Bonus Meta(" & Format(w_Piso, "0.00") & " * " & Format(wPercMeta, "0.0") & "% = " & Format(vrMetaBonus, "0.00") & ") "
                        vrPisoMeta = w_Piso + vrMetaBonus
                     Else
                        vrPisoMeta = w_Piso
                        descMeta = ""
                     End If
                 
                Else
                 
                     If vrVenda >= vrMeta Then
                        'Bonus de 15% do Salario
                        vrMetaBonus = w_Piso * 0.15
                        'descMeta = " + Bonus Meta(" & vrSalario & " * 15% = " & vrMetaBonus & ") "
                        descMeta = " + Bonus Meta(" & Format(w_Piso, "0.00") & " * 15% = " & Format(vrMetaBonus, "0.00") & ") "
                        vrPisoMeta = w_Piso + vrMetaBonus
                     Else
                        vrPisoMeta = w_Piso
                        descMeta = ""
                     End If
                    '*****
                
                End If
             End If
             

                 
             If (w_DtDiff < w_ultDiaMes) And (w_DtDiff < 30) Then
                 de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 22, "+", Format(vrPisoMeta, "0.00"), "#FIXO PROPORCIONAL# Fixo: " & Format(w_PisoOriginal, "0.00") & " / " & CInt(w_ultDiaMes) & " = " & Format(w_PisoOriginal / w_ultDiaMes, "0.00") & " * " & w_DtDiff & " dias = " & Format(w_Piso, "0.00") & descMeta & " [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
             Else
                 de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFICHA"), 22, "+", Format(vrPisoMeta, "0.00"), "#FIXO# Fixo: " & Format(w_Piso, "0.00") & descMeta & " [GERADO AUTOMATICAMENTE]", ADOREG.Recordset.Fields("M_LOGO"), "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
             End If
                 
              
       End If
       ADOREG.Recordset.MoveNext
       
       If ADOREG.Recordset.EOF And cbMostrar.text = "CX EXTRA" Then
            cbMostrar.text = "2º CAIXA"
            cmdMostrar_Click
            ADOREG.Recordset.MoveFirst
       End If
            
    Loop
    
    ADOREG.Recordset.MoveFirst
    adoComis.Close
    de.rscmdComissGerar.Close
    Lancamentos

End Sub

Private Sub cmdDelFichas_Click()
On Error GoTo err1

   Op_Click 5
    
    ADOREG.Recordset.MoveFirst
    Do While Not ADOREG.Recordset.EOF
    
       'PEGA DT DE DEMISSÃO
        w_dt_DEM = IIf(TXT_DT_DEM = "", Format(txt_DT_ADM, "DD/MM/YYYY"), Format(TXT_DT_DEM, "DD/MM/YYYY"))
       'Se não tiver fichas abertas
       If de.cnc.Execute("SELECT M_NFICHA FROM TAB_FICHA_MENS WHERE M_F_COD = " & txt_F_COD & " AND iif(M_MES>01,(M_MES = " & CInt(TXT_MES) - 1 & " AND M_ANO = " & TXT_ANO & "),(M_MES = 12 AND M_ANO = " & CInt(TXT_ANO) - 1 & "))").RecordCount = 0 Then
            'Se for admissão no mês atual
            If Format(txt_DT_ADM, "YYYY") = Format(TXT_ANO, "0000") And Format(txt_DT_ADM, "MM") = Format(TXT_MES, "00") Then
                'msgbox
            Else
            
                 w_logo = txtLogo2
                 w_Nome = TXT_FUNC
            
                 'ATUALIZA DATA DE DEMISSÃO NO CADASTRO PRA NÃO GERAR MAIS
                 de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_DT_DEM = '" & w_dt_DEM & "'  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
                 'DELETA FICHA ATUAL
                 de.cnc.Execute "DELETE * FROM TAB_FICHA_MENS WHERE M_NFICHA = " & TXT_NFICHA & "", REG_AF
                 If REG_AF = 1 Then
                     'MsgBox "Funcionário " & w_nome & " (" & w_logo & ") excluído com sucesso do mês " & TXT_MES & "/" & TXT_ANO & ".", vbInformation
                 Else
                     MsgBox "Funcionário " & w_Nome & " (" & w_logo & ") NÃO EXCLUÍDO do mês " & TXT_MES & "/" & TXT_ANO & ".", vbExclamation
                 End If
            End If
           
       End If
    
       ADOREG.Recordset.MoveNext
       
    Loop
    
    cmdAtualizar_Click

    w_Op = -Op(1).value * 1 - Op(2).value * 2 - Op(3).value * 3 - Op(4).value * 4 - Op(5).value * 5 - Op(6).value * 6 - Op(7).value * 7
    Op_Click (w_Op)
    cmdFiltrar_Click

sair:
    Exit Sub

err1:
    
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    
    Resume sair
End Sub

Private Sub cmdDesbloquear_Click()
    
    If Trim(TXT_AC_F) = "" Then
        de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_DT_ACF = NULL WHERE (M_NFICHA = " & TXT_NFICHA & " )", w_reg
        ADOREG.Recordset.Fields("M_DT_ACF") = Null
    End If
    
    If IsNull(ADOREG.Recordset.Fields("M_DT_ACF")) Then
        If ADOREG.Recordset.Fields("M_BLOQ") Then
            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_BLOQ = " & 0 & " Where (M_nficha = " & TXT_NFICHA & ")"
            de.cmdIncluirLog Date, Time, w_usuario, "EDITAR", "FICHA", "FICHA: " & TXT_NFICHA & " | FUNCIONÁRIO: " & TXT_FUNC & " | ## LIBERANDO ##"
        Else
            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_BLOQ = " & -1 & " Where (M_nficha = " & TXT_NFICHA & ")"
            de.cmdIncluirLog Date, Time, w_usuario, "EDITAR", "FICHA", "FICHA: " & TXT_NFICHA & " | FUNCIONÁRIO: " & TXT_FUNC & " | ## BLOQUEANDO ##"
        End If
        
        Cancelar
        Editar 0
        
        If ADOREG.Recordset.Fields("M_BLOQ") = False Then BarraF.Buttons("editar").Enabled = True
        
    Else
          MsgBox "Não é possível desbloquear uma ficha com CARIMBO!", vbCritical
    End If
   
End Sub

Private Sub cmdDescCalcFixo_Click()
On Error GoTo err1

Dim adoFixos As ADODB.Recordset

Set adoFixos = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_EMP_COD = " & txt_F_COD).Clone

Do While Not adoFixos.EOF
    de.cnc.Execute ("DELETE FROM TAB_DESC_CALC Where C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFicha") & " AND C_NCRED = " & adoFixos.Fields("CF_CODIGO"))
    de.cmdIncluirDescCalc2 Date, ADOREG.Recordset.Fields("M_NFicha"), adoFixos.Fields("CF_TP_CONTA"), adoFixos.Fields("CF_TP_OP"), adoFixos.Fields("CF_VALOR"), adoFixos.Fields("CF_DESC"), "0", adoFixos.Fields("CF_CODIGO"), "0", "0", adoFixos.Fields("CF_EMP_COD"), 0
    adoFixos.MoveNext
Loop

'de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFicha"), 20, "+", vrSalario, wDesc, "", "0", "0", "0", "0"
'End If
Cancelar
Editar 0

sair:
    Exit Sub
err1:
    Resume sair

End Sub

Private Sub cmdDespesa_Click()
Dim dtIni, dtFim As Date

    frm_ESCOLHA_DATA.Show 1
    
    dtIni = CDate(frm_ESCOLHA_DATA.TXT_DT_INICIAL)
    dtFim = CDate(frm_ESCOLHA_DATA.TXT_DT_FINAL)
   
    'If de.rscmdComiss_Grouping.State = 1 Then de.rscmdComiss_Grouping.Close
    
    On Error Resume Next
    'de.cmdDROPtmpDespesa1
    'de.cmdDROPtmpDespesa2
    'de.cmdDROPtmpDespesa3
    
    'de.cmdCREATEtmpDespesa1
    'de.cmdCREATEtmpDespesa2
    'de.cmdCREATEtmpDespesa3
    
    de.cmdDELETEtmpDespesa1
    de.cmdDELETEtmpDespesa2
    de.cmdDELETEtmpDespesa3
    
    de.cmdAddtmpDespesa1 dtIni, dtFim
    de.cmdAddtmpDespesa2
    de.cmdAddtmpDespesa3 dtIni, dtFim
    
    frm_Despesas.Show
    
    frm_Despesas.txtData = dtIni
        
End Sub

Private Sub cmdDespesa99_Click()
    If MsgBox("Gerar Despesas 92 p/ o mês " & lblMes.Caption & "?", vbYesNo, "Despesas 92") = vbYes Then
    Dim wNumContas
        wNumContas = 99
    
        'Despesas RP
        Dim totalRP, val As Double
        Dim adoRP As ADODB.Recordset
        Set adoRP = de.cnc.Execute("SELECT TAB_VENDA.V_F_LOJA as NUM, TAB_VENDA.V_VR, (SELECT SUM(TAB_VENDA.V_VR) FROM TAB_VENDA " & _
                                    "WHERE (((Month([V_DATA]))=" & TXT_MES & ") AND ((Year([V_DATA]))=" & TXT_ANO & "));) AS TOTAL, [V_VR]/[TOTAL] AS PERC " & _
                                    "FROM TAB_VENDA WHERE (((Month([V_DATA]))=" & TXT_MES & " ) AND ((Year([V_DATA]))=" & TXT_ANO & "));").Clone
        totalRP = InputBox("Digite o valor de TOTAL FIXO da 92 no mês " & TXT_MES & "/" & TXT_ANO & ":", "Total fixo da 92")
        Dim qtRP As Integer
        qtRP = 0
        If de.cncContas.State = 0 Then de.cncContas.Open
        Do While Not adoRP.EOF
            val = 0
            val = totalRP * adoRP.Fields("PERC")
            wNumContas = de.cncContas.Execute("SELECT tblloja_loja_contas FROM tblloja where tblloja_loja_fichas = '" & Format(adoRP.Fields("NUM"), "000") & "'").Fields(0)
            de.cmdAddProcesso "1603", CDate(Now()), "**Incluído automaticamente**", "R", CStr(val), wNumContas, CDate(Now())
            lastNumProcesso = de.cncContas.Execute("SELECT MAX(tblpro_num) FROM tblpropg").Fields(0)
            de.cmdAddProcessoItem lastNumProcesso, CDate("01/" & TXT_MES & "/" & TXT_ANO), CStr(val), CDate("01/" & TXT_MES & "/" & TXT_ANO)
            qtRP = qtRP + 1
            adoRP.MoveNext
        Loop
        MsgBox "Foram incluídos " & qtRP & " processos de Despesas RP no sistema de Contas, na data de " & CDate("01/" & TXT_MES & "/" & TXT_ANO) & ".", vbInformation, "Despesas RP"
        'FIM Despesas RP
    End If
End Sub

Private Sub cmdDtDEM_Click()
On Error GoTo err1

    Op_Click 5
    
    ADOREG.Recordset.MoveFirst
    Do While Not ADOREG.Recordset.EOF
    
       'ATUALIZA DATA DE DEMISSÃO DO CADASTRO
        w_dt_DEM = IIf(TXT_DT_DEM = "", Null, Format(TXT_DT_DEM, "DD/MM/YYYY"))
        'Se não tiver fichas abertas
        If de.cnc.Execute("SELECT M_NFICHA FROM TAB_FICHA_MENS WHERE M_DT_DEM IS NULL AND M_F_COD = " & txt_F_COD & " AND ((M_MES >= " & TXT_MES & " AND M_ANO = " & TXT_ANO & ") OR M_ANO > " & TXT_ANO & ")").RecordCount = 0 Then
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_DT_DEM = '" & w_dt_DEM & "'  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        Else
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_DT_DEM = NULL  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        End If
        
        
        'REMOVE VISTO SE NÃO TIVER DATA DE DEMISSÃO
        If ADOREG.Recordset.Fields("M_DEM_OK") Then
            If TXT_DT_DEM = "" Then
                de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_DEM_OK = 0  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
                de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_DEM_OK = 0  WHERE (M_NFICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & " )", w_reg
            End If
        End If
        
         ADOREG.Recordset.MoveNext
    Loop

    w_Op = -Op(1).value * 1 - Op(2).value * 2 - Op(3).value * 3 - Op(4).value * 4 - Op(5).value * 5 - Op(6).value * 6 - Op(7).value * 7
    Op_Click (w_Op)
    cmdFiltrar_Click

sair:
    Exit Sub

err1:
    
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    
    Resume sair
End Sub

Private Sub cmdEmprestimo_Click()
    '*****  PRESTAÇÕES DE EMRPESTIMO ****
    If de.rsTAB_FICHA_MENS.State = 1 Then de.rsTAB_FICHA_MENS.Close
    de.TAB_FICHA_MENS
    
    Dim W_ADO_EMP As ADODB.Recordset
    'Zera a descrição dos q/ tem saldo zero
    Set W_ADO_EMP = de.cnc.Execute("SELECT * FROM TAB_EMPRESTIMO, TAB_Funcionario WHERE E_SALDO = 0 and TAB_EMPRESTIMO.E_F_COD = TAB_Funcionario.F_CODIGO AND TAB_Funcionario.F_COD_L LIKE '" & TXT_LOGO & "%'").Clone
    Do While Not W_ADO_EMP.EOF
        '*** Dá baixa no emprestimo na tab. funcionario ***
        de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_EMPRESTIMO_ANOT = '' WHERE (F_Codigo = " & W_ADO_EMP.Fields("E_F_COD") & ")"
        W_ADO_EMP.MoveNext
    Loop
    
    Set W_ADO_EMP = de.cnc.Execute("SELECT * FROM TAB_EMPRESTIMO, TAB_Funcionario WHERE E_SALDO > 0 and TAB_EMPRESTIMO.E_F_COD = TAB_Funcionario.F_CODIGO AND TAB_Funcionario.F_COD_L LIKE '" & TXT_LOGO & "%'").Clone
    
    Do While Not W_ADO_EMP.EOF
        '*** CALCULA SOMENTE SE EXISTIR FICHA NESTE MÊS -****
        If de.cnc.Execute("SELECT M_NFICHA FROM TAB_FICHA_MENS WHERE M_F_COD = " & W_ADO_EMP.Fields("E_F_COD") & " AND M_MES = " & TXT_MES & " AND M_ANO = " & TXT_ANO & "").RecordCount > 0 Then
                w_parc = de.cnc.Execute("Select EP_Parc from tab_Emprestimo_pg Where ep_codigo = " & W_ADO_EMP.Fields("E_codigo") & " and ep_parc > 0 ").RecordCount + 1
                W_DT_PG = CVDate("01/" & TXT_MES & "/" & TXT_ANO) + 32
                
                If IsDate((W_ADO_EMP.Fields("E_DIA_PG") & "/" & TXT_MES & "/" & TXT_ANO)) Then
                    W_DT_PG = CVDate(W_ADO_EMP.Fields("E_DIA_PG") & "/" & TXT_MES & "/" & TXT_ANO) + 31
                    W_DT_PG = CVDate(W_ADO_EMP.Fields("E_DIA_PG") & "/" & Format(W_DT_PG, "mm/yyyy"))
                Else
                    W_DT_PG = CVDate("01/" & TXT_MES & "/" & TXT_ANO) + 32
                    W_DT_PG = CVDate("01/" & Format(W_DT_PG, "mm/yyyy")) - 1
                    If CDbl(Format(W_DT_PG, "dd")) < W_ADO_EMP.Fields("E_DIA_PG") Then
                        w_QtDias = W_ADO_EMP.Fields("E_DIA_PG") - CDbl(Format(W_DT_PG, "dd"))
                    End If
                    W_DT_PG = W_DT_PG + w_QtDias
                End If
                
                
                W_JUROS = Format(CALC_PG_EMP(W_ADO_EMP, W_DT_PG), "R$ 0.00")
                w_Valor = (W_ADO_EMP.Fields("E_SALDO") / IIf(W_PARC_RESTANTE = 0, 1, W_PARC_RESTANTE)) + CDbl(W_JUROS)
                                        
                W_NFICHA = de.cnc.Execute("SELECT M_NFICHA FROM TAB_FICHA_MENS WHERE M_F_COD = " & W_ADO_EMP.Fields("E_F_COD") & " AND M_MES = " & TXT_MES & " AND M_ANO = " & TXT_ANO & "").Fields(0)
                
                W_DESC_CONTA = "Pg. Emp.: " & W_ADO_EMP.Fields("E_QT_PG") + 1 & "/" & W_ADO_EMP.Fields("E_QT_PARC") & vbCrLf & "Valor : " & Format(w_Valor - W_JUROS, "R$ 0.00") & "    +    Juros : " & Format(W_JUROS, "R$ 0.00")
                
                '*** INCLUI CONTA P/ DESCONTO DO EMP. ***
                de.cmdIncluirDescCalc W_DT_PG, W_NFICHA, "9", "-", CDbl(w_Valor * -1), W_DESC_CONTA, "0", "0", CDbl(W_JUROS), w_parc, W_ADO_EMP.Fields("E_CODIGO")
                '*** iNCLUINDO PAGAMENTO DE EMPRESTIMO  -  TAB_EMPRESTIMO_PG ***
                W_C_COD = de.cnc.Execute("SELECT MAX(C_CODIGO)AS COD FROM TAB_DESC_CALC WHERE C_N_FICHA = " & W_NFICHA & "").Fields(0)
                de.cmdIncluirEmprestimoPG W_ADO_EMP.Fields("E_CODIGO"), W_DT_PG, w_parc, w_qt_dias, CDbl(CDbl(w_Valor) - CDbl(W_JUROS)), CDbl(W_JUROS), W_C_COD
    
                '*** Dá baixa no emprestimo na tab. funcionario ***
                de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_EMPRESTIMO = F_EMPRESTIMO - ' " & CDbl(w_Valor - W_JUROS) & "' WHERE (F_Codigo = " & W_ADO_EMP.Fields("E_F_COD") & ")"
                
                '*** Dá baixa no emprestimo na tab. emprestimo ***
                de.cnc.Execute "UPDATE TAB_EMPRESTIMO SET E_QT_PG = E_QT_PG + 1 , E_DT_ULT_PG = '" & W_DT_PG & "', E_SALDO = E_SALDO - '" & CDbl(w_Valor - W_JUROS) & "' WHERE (E_Codigo = " & W_ADO_EMP.Fields("E_CODIGO") & ")"
                
                '*** ATUALIZAR A ANOTAÇÃO DO EMPRESTIMO DO FUNCIONARIO ***
                    '** Sql EMP. P/ GRID
                        
                        W_EMP_ANOT = ""
                        Dim ADO_ANOT As ADODB.Recordset
                        
                        w_Dt = CVDate("01/" & TXT_MES & "/" & TXT_ANO) + 65
                        Set ADO_ANOT = de.cnc.Execute("SELECT * FROM TAB_EMPRESTIMO WHERE E_F_COD = " & W_ADO_EMP.Fields("E_F_COD") & " AND (E_SALDO > 0  OR E_DT_ULT_PG <= #" & Format(w_Dt, "MM/DD/YYYY") & "#)").Clone
                        Do While Not ADO_ANOT.EOF
                            W_EMP_ANOT = W_EMP_ANOT & IIf(Len(W_EMP_ANOT) > 0, vbCrLf, "") & ". Dt Emp.: " & ADO_ANOT.Fields("E_DT_EMP") & "    Valor Emp.: " & Format(ADO_ANOT.Fields("E_VALOR"), "R$ 0.00") & "     Juros : " & ADO_ANOT.Fields("E_Juro_ao_mes") * 100 & " %" & "     Parc. Pg.: " & ADO_ANOT.Fields("E_QT_PG") & " / " & ADO_ANOT.Fields("E_QT_PARC")
                            W_EMP_ANOT = W_EMP_ANOT & vbCrLf & ". Saldo Ant.: " & Format(W_ADO_EMP.Fields("E_SALDO"), "R$ 0.00") & "         Dt Ult. Pg.: " & ADO_ANOT.Fields("E_DT_ULT_PG") & "        Saldo At.: " & Format(ADO_ANOT.Fields("E_SALDO"), "R$ 0.00")
                        
                            ADO_ANOT.MoveNext
                        Loop
                        
                        '*** UPDATE NO FUNCIONARIO ATUALIZANDO A ANOTAÇÃO DO EMPRESTIMO ***
                        de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_EMPRESTIMO_ANOT = '" & W_EMP_ANOT & "' WHERE (F_Codigo = " & W_ADO_EMP.Fields("E_F_COD") & ")"
            
                        '*** Atualiza o Valor Total da Ficha ***
                        de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_TOTAL = M_TOTAL + '" & w_Valor & "', M_EMPRESTIMO_ANOT = '" & IIf(W_EMP_ANOT = "", " ", W_EMP_ANOT) & "' WHERE (M_NFICHA = " & W_NFICHA & ")"
    
        End If
    
        W_ADO_EMP.MoveNext
    Loop
    
    Set W_ADO_EMP = Nothing
End Sub

Private Sub cmdEsconder_Click()
    mnuAcessoTotal_Click
End Sub



Private Sub cmdExportarContas_Click()
On Error GoTo err1
    '***Exportação para Sistema de Contas a Pagar
    Dim lastNumProcesso
    Dim wNumContas
        wNumContas = 99
        
    
    
    If MsgBox("Gerar GASOLINA p/ o mês " & lblMes.Caption & "?", vbYesNo, "GERAR GASOLINA") = vbYes Then
        'GASOLINA
        Dim adoGasolina As ADODB.Recordset
        Set adoGasolina = de.cnc.Execute("SELECT Lojb010.NUM, Sum(TAB_DESC_CALC.C_VALOR) AS VAL FROM Lojb010 INNER JOIN (TAB_FICHA_MENS INNER JOIN TAB_DESC_CALC ON TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA) ON Lojb010.COD_LOJ = TAB_FICHA_MENS.M_LOGO WHERE (((TAB_FICHA_MENS.M_ANO)= " & TXT_ANO & ")" & _
                                         "  AND ((TAB_FICHA_MENS.M_MES)= " & TXT_MES & ") AND ((TAB_DESC_CALC.C_TP_CONTA)=158 Or (TAB_DESC_CALC.C_TP_CONTA)=162)) GROUP BY Lojb010.NUM;").Clone
        
        
        Dim qtGasolina As Integer
        qtGasolina = 0
        

        If de.cncContas.State = 0 Then de.cncContas.Open
        Do While Not adoGasolina.EOF
        'If adoGasolina.Fields("NUM") = "99" Then
            wNumContas = de.cncContas.Execute("SELECT tblloja_loja_contas FROM tblloja where tblloja_loja_fichas = '" & Format(adoGasolina.Fields("NUM"), "000") & "'").Fields(0)
            de.cmdAddProcesso "1486", CDate(Now()), "**Incluído automaticamente**", "R", CStr(adoGasolina.Fields("VAL")), wNumContas, CDate(Now())
            lastNumProcesso = de.cncContas.Execute("SELECT MAX(tblpro_num) FROM tblpropg").Fields(0)
            de.cmdAddProcessoItem lastNumProcesso, CDate("01/" & TXT_MES & "/" & TXT_ANO), CStr(adoGasolina.Fields("VAL")), CDate("01/" & TXT_MES & "/" & TXT_ANO)
            qtGasolina = qtGasolina + 1
        'End If
            adoGasolina.MoveNext
        Loop
        
        MsgBox "Foram incluídos " & qtGasolina & " processos de Gasolina no sistema de Contas, na data de " & CDate("01/" & TXT_MES & "/" & TXT_ANO) & ".", vbInformation, "Exportação de Gasolina"
        'FIM GASOLINA
    End If
        
    If MsgBox("Gerar SALÁRIO p/ o mês " & lblMes.Caption & "?", vbYesNo, "GERAR SALÁRIO") = vbYes Then
        'SALARIO
        Dim adoSalario As ADODB.Recordset
        Set adoSalario = de.cnc.Execute("SELECT Lojb010.NUM, Sum(TAB_DESC_CALC.C_VALOR) AS VAL FROM Lojb010 INNER JOIN (TAB_FICHA_MENS INNER JOIN TAB_DESC_CALC ON TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA) ON Lojb010.COD_LOJ = TAB_FICHA_MENS.M_LOGO " & _
                                        " GROUP BY Lojb010.NUM, TAB_FICHA_MENS.M_ANO, TAB_FICHA_MENS.M_MES, TAB_DESC_CALC.C_TP_OP " & _
                                        " HAVING (((TAB_FICHA_MENS.M_ANO)=" & TXT_ANO & ") AND ((TAB_FICHA_MENS.M_MES)=" & TXT_MES & ") AND ((TAB_DESC_CALC.C_TP_OP)='+'))").Clone
        
        Dim qtSalario As Integer
        qtSalario = 0
        
        If de.cncContas.State = 0 Then de.cncContas.Open
        Do While Not adoSalario.EOF
        If adoSalario.Fields("NUM") <> "98" Then 'não pega salários da RP
            wNumContas = de.cncContas.Execute("SELECT tblloja_loja_contas FROM tblloja where tblloja_loja_fichas = '" & Format(adoSalario.Fields("NUM"), "000") & "'").Fields(0)
            de.cmdAddProcesso "1487", CDate(Now()), "**Incluído automaticamente**", "R", CStr(adoSalario.Fields("VAL")), wNumContas, CDate(Now())
            lastNumProcesso = de.cncContas.Execute("SELECT MAX(tblpro_num) FROM tblpropg").Fields(0)
            de.cmdAddProcessoItem lastNumProcesso, CDate("01/" & TXT_MES & "/" & TXT_ANO), CStr(adoSalario.Fields("VAL")), CDate("01/" & TXT_MES & "/" & TXT_ANO)
            qtSalario = qtSalario + 1
        End If
            adoSalario.MoveNext
        Loop

        MsgBox "Foram incluídos " & qtSalario & " processos de Salário no sistema de Contas, na data de " & CDate("01/" & TXT_MES & "/" & TXT_ANO) & ".", vbInformation, "Exportação de Salário"
        'FIM SALARIO
    End If
    
    If MsgBox("Gerar PAGTO PASSE - VT p/ o mês " & lblMes.Caption & "?", vbYesNo, "GERAR PAGTO PASSE - VT") = vbYes Then
        'PAGTO PASSE - VALE TRANSPORTE
        Dim adoPasse As ADODB.Recordset
        Set adoPasse = de.cnc.Execute("SELECT Lojb010.NUM, Sum(TAB_DESC_CALC.C_VALOR) AS VAL FROM Lojb010 INNER JOIN (TAB_FICHA_MENS INNER JOIN TAB_DESC_CALC ON TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA) ON Lojb010.COD_LOJ = TAB_FICHA_MENS.M_LOGO " & _
                                         " GROUP BY Lojb010.NUM, TAB_FICHA_MENS.M_ANO, TAB_FICHA_MENS.M_MES, TAB_DESC_CALC.C_TP_CONTA " & _
                                         " HAVING (((TAB_FICHA_MENS.M_ANO)= " & TXT_ANO & ") AND ((TAB_FICHA_MENS.M_MES)= " & TXT_MES & ") AND ((TAB_DESC_CALC.C_TP_CONTA)=111))").Clone
        
        Dim qtPasse As Integer
        qtPasse = 0
        
        If de.cncContas.State = 0 Then de.cncContas.Open
        Do While Not adoPasse.EOF
        'If adoPasse.Fields("NUM") = "99" Then
            wNumContas = de.cncContas.Execute("SELECT tblloja_loja_contas FROM tblloja where tblloja_loja_fichas = '" & Format(adoPasse.Fields("NUM"), "000") & "'").Fields(0)
            de.cmdAddProcesso "1594", CDate(Now()), "**Incluído automaticamente**", "R", CStr(adoPasse.Fields("VAL")), wNumContas, CDate(Now())
            lastNumProcesso = de.cncContas.Execute("SELECT MAX(tblpro_num) FROM tblpropg").Fields(0)
            de.cmdAddProcessoItem lastNumProcesso, CDate("01/" & TXT_MES & "/" & TXT_ANO), CStr(adoPasse.Fields("VAL")), CDate("01/" & TXT_MES & "/" & TXT_ANO)
            qtPasse = qtPasse + 1
        'End If
            adoPasse.MoveNext
        Loop
        
        MsgBox "Foram incluídos " & qtPasse & " processos de Pagto. de Passe - VT no sistema de Contas, na data de " & CDate("01/" & TXT_MES & "/" & TXT_ANO) & ".", vbInformation, "Exportação de Pagamento de Passes - VT"
        'FIM PAGTO PASSE - VALE TRANSPORTE
    End If
        
    If MsgBox("Gerar ABS p/ o mês " & lblMes.Caption & "?", vbYesNo, "GERAR ABS") = vbYes Then
        'ABS
        Dim adoABS As ADODB.Recordset
        If de.cncContas.State = 0 Then de.cncContas.Open
        Set adoABS = de.cncContas.Execute("SELECT tblloja_loja_contas as NUM FROM tblloja WHERE (tblloja_loja_ordem Is Not Null)").Clone
        
        Dim qtABS As Integer
        qtABS = 0

        Do While Not adoABS.EOF
        'If adoABS.Fields("NUM") = "99" Then
            wNumContas = de.cncContas.Execute("SELECT tblloja_loja_contas FROM tblloja where tblloja_loja_fichas = '" & Format(adoABS.Fields("NUM"), "000") & "'").Fields(0)
            de.cmdAddProcesso "473", CDate(Now()), "**Incluído automaticamente** R$414,00 mensalidade + 1/12 do 13º (R$414,00 / 12 = R$34,50) = R$448,50", "R", "448,50", wNumContas, CDate(Now())
            lastNumProcesso = de.cncContas.Execute("SELECT MAX(tblpro_num) FROM tblpropg").Fields(0)
            de.cmdAddProcessoItem lastNumProcesso, CDate("01/" & TXT_MES & "/" & TXT_ANO), "448,50", CDate("01/" & TXT_MES & "/" & TXT_ANO)
            qtABS = qtABS + 1
        'End If
            adoABS.MoveNext
        Loop
        
        MsgBox "Foram incluídos " & qtABS & " processos de ABS no sistema de Contas, na data de " & CDate("01/" & TXT_MES & "/" & TXT_ANO) & ".", vbInformation, "Exportação de Mensalidade ABS"
        'FIM ABS
    End If
    
    If MsgBox("Gerar Despesas Retaguarda p/ o mês " & lblMes.Caption & "?", vbYesNo, "Despesas Retaguarda") = vbYes Then
        'Despesas RP
        Dim totalRP, val As Double
        Dim adoRP As ADODB.Recordset
        Set adoRP = de.cnc.Execute("SELECT TAB_VENDA.V_F_LOJA as NUM, TAB_VENDA.V_VR, (SELECT SUM(TAB_VENDA.V_VR) FROM TAB_VENDA " & _
                                    "WHERE (((Month([V_DATA]))=" & TXT_MES & ") AND ((Year([V_DATA]))=" & TXT_ANO & "));) AS TOTAL, [V_VR]/[TOTAL] AS PERC " & _
                                    "FROM TAB_VENDA WHERE (((Month([V_DATA]))=" & TXT_MES & " ) AND ((Year([V_DATA]))=" & TXT_ANO & "));").Clone
        totalRP = 180000
        Dim qtRP As Integer
        qtRP = 0

        If de.cncContas.State = 0 Then de.cncContas.Open
        Do While Not adoRP.EOF
        'If adoRP.Fields("NUM") = "99" Then
            val = 0
            val = totalRP * adoRP.Fields("PERC")
            wNumContas = de.cncContas.Execute("SELECT tblloja_loja_contas FROM tblloja where tblloja_loja_fichas = '" & Format(adoRP.Fields("NUM"), "000") & "'").Fields(0)
            de.cmdAddProcesso "1598", CDate(Now()), "**Incluído automaticamente**", "R", CStr(val), wNumContas, CDate(Now())
            lastNumProcesso = de.cncContas.Execute("SELECT MAX(tblpro_num) FROM tblpropg").Fields(0)
            de.cmdAddProcessoItem lastNumProcesso, CDate("01/" & TXT_MES & "/" & TXT_ANO), CStr(val), CDate("01/" & TXT_MES & "/" & TXT_ANO)
            qtRP = qtRP + 1
        'End If
            adoRP.MoveNext
        Loop
        
        MsgBox "Foram incluídos " & qtRP & " processos de Despesas Retaguarda no sistema de Contas, na data de " & CDate("01/" & TXT_MES & "/" & TXT_ANO) & ".", vbInformation, "Exportação de Despesas RP"
        'FIM Despesas RP
    End If
        
    If MsgBox("Gerar Taxa de Retenção dos Cartões p/ o mês " & lblMes.Caption & "?", vbYesNo, "Taxa de Retenção dos Cartões") = vbYes Then
        'Taxa Retenção Cartão
        
        Dim adoCartao As ADODB.Recordset
        If de.cncSisInt.State = 0 Then de.cncSisInt.Open
        de.cmdTaxaCartao TXT_MES, TXT_ANO
        Set adoCartao = de.rscmdTaxaCartao.Clone

        Dim qtCartao As Integer
        qtCartao = 0
        
        
        If de.cncContas.State = 0 Then de.cncContas.Open
        Do While Not adoCartao.EOF
        'If adoCartao.Fields("logo") = "99" Then
            wNumContas = de.cncContas.Execute("SELECT tblloja_loja_contas FROM tblloja where tblloja_loja_fichas = '" & Format(adoCartao.Fields("logo"), "000") & "'").Fields(0)
            de.cmdAddProcesso "1608", CDate(Now()), "**Incluído automaticamente**", "R", CStr(CDbl(adoCartao.Fields("taxaCartao"))), wNumContas, CDate(Now())
            lastNumProcesso = de.cncContas.Execute("SELECT MAX(tblpro_num) FROM tblpropg").Fields(0)
            de.cmdAddProcessoItem lastNumProcesso, CDate("01/" & TXT_MES & "/" & TXT_ANO), CStr(CDbl(adoCartao.Fields("taxaCartao"))), CDate("01/" & TXT_MES & "/" & TXT_ANO)
            qtCartao = qtCartao + 1
        'End If
            adoCartao.MoveNext
        Loop

        MsgBox "Foram incluídos " & qtCartao & " processos de Taxa de Retenção no sistema de Contas, na data de " & CDate("01/" & TXT_MES & "/" & TXT_ANO) & ".", vbInformation, "Exportação de Taxa de Retenção de Cartão"
        'FIM Taxa Retenção Cartão
    End If
        
        
    If MsgBox("Gerar Mensalidade das Maquinetas p/ o mês " & lblMes.Caption & "?", vbYesNo, "Mensalidade das Maquinetas") = vbYes Then
        'Mensalidade das Maquinetas
        
'        Public Type bVal
'            b As String
'            val As Double
'        End Type

        Dim qtMaquinetas As Integer
        qtMaquinetas = 0
        
        Dim maquinetas(41) As bVal
                
'Atualizado em 17/03/2016
        maquinetas(0).b = "00"
        maquinetas(0).val = 264.7 + 29.61
        
        maquinetas(1).b = "01"
        maquinetas(1).val = 193.7 + 29.61
        
        maquinetas(2).b = "03"
        maquinetas(2).val = 251.8 + 29.61
        
        maquinetas(3).b = "04"
        maquinetas(3).val = 251.8 + 29.61
       
        maquinetas(4).b = "06"
        maquinetas(4).val = 264.7 + 29.61
        
        maquinetas(5).b = "08"
        maquinetas(5).val = 251.8 + 29.61
        
        maquinetas(6).b = "09"
        maquinetas(6).val = 251.8 + 29.61
        
        maquinetas(7).b = "10"
        maquinetas(7).val = 251.8 + 29.61
        
        maquinetas(8).b = "11"
        maquinetas(8).val = 251.8 + 29.61
        
        maquinetas(9).b = "12"
        maquinetas(9).val = 251.8 + 29.61
        
        maquinetas(10).b = "13"
        maquinetas(10).val = 251.8 + 29.61
        
        maquinetas(11).b = "14"
        maquinetas(11).val = 251.8 + 29.61
        
        maquinetas(12).b = "16"
        maquinetas(12).val = 251.8 + 29.61
        
        maquinetas(13).b = "17"
        maquinetas(13).val = 264.7 + 29.61
        
        maquinetas(14).b = "18"
        maquinetas(14).val = 359.7 + 29.61
        
        maquinetas(15).b = "19"
        maquinetas(15).val = 264.7 + 29.61
        
        maquinetas(16).b = "89"
        maquinetas(16).val = 264.7 + 29.61
        
        maquinetas(17).b = "33"
        maquinetas(17).val = 251.8 + 29.61
        
        maquinetas(18).b = "34"
        maquinetas(18).val = 231.8 + 29.61
        
        maquinetas(19).b = "35"
        maquinetas(19).val = 191.9 + 29.61
        
        maquinetas(20).b = "40"
        maquinetas(20).val = 251.8 + 29.61
        
        maquinetas(21).b = "45"
        maquinetas(21).val = 264.7 + 29.61
        
        maquinetas(22).b = "50"
        maquinetas(22).val = 180.8 + 29.61
        
        maquinetas(23).b = "51"
        maquinetas(23).val = 251.8 + 29.61
        
        maquinetas(24).b = "53"
        maquinetas(24).val = 251.8 + 29.61
        
        maquinetas(25).b = "54"
        maquinetas(25).val = 264.7 + 29.61

        maquinetas(26).b = "57"
        maquinetas(26).val = 155.6 + 29.61
        
        maquinetas(27).b = "58"
        maquinetas(27).val = 180.8 + 29.61
        
        maquinetas(28).b = "59"
        maquinetas(28).val = 251.8 + 29.61
        
        maquinetas(29).b = "60"
        maquinetas(29).val = 251.8 + 29.61
        
        maquinetas(30).b = "61"
        maquinetas(30).val = 251.8 + 29.61
        
        maquinetas(31).b = "62"
        maquinetas(31).val = 193.7 + 29.61
        
        maquinetas(32).b = "63"
        maquinetas(32).val = 264.7 + 29.61
        
        maquinetas(33).b = "64"
        maquinetas(33).val = 264.7 + 29.61
        
        maquinetas(34).b = "65"
        maquinetas(34).val = 251.8 + 29.61
        
        maquinetas(35).b = "66"
        maquinetas(35).val = 251.8 + 29.61
        
        maquinetas(36).b = "67"
        maquinetas(36).val = 264.7 + 29.61
        
        maquinetas(37).b = "68"
        maquinetas(37).val = 251.8 + 29.61
        
        maquinetas(38).b = "69"
        maquinetas(38).val = 193.7 + 29.61
        
        maquinetas(39).b = "70"
        maquinetas(39).val = 264.7 + 29.61
        
        maquinetas(40).b = "71"
        maquinetas(40).val = 193.7 + 29.61
       
       If de.cncContas.State = 0 Then de.cncContas.Open
       For I = 0 To 40
         'If adoCartao.Fields("logo") = "99" Then
             wNumContas = de.cncContas.Execute("SELECT tblloja_loja_contas FROM tblloja where tblloja_loja_fichas = '" & Format(maquinetas(I).b, "000") & "'").Fields(0)
             de.cmdAddProcesso "1609", CDate(Now()), "**Incluído automaticamente**", "R", CStr(CDbl(maquinetas(I).val)), wNumContas, CDate(Now())
             lastNumProcesso = de.cncContas.Execute("SELECT MAX(tblpro_num) FROM tblpropg").Fields(0)
             de.cmdAddProcessoItem lastNumProcesso, CDate("01/" & TXT_MES & "/" & TXT_ANO), CStr(CDbl(maquinetas(I).val)), CDate("01/" & TXT_MES & "/" & TXT_ANO)
             qtMaquinetas = qtMaquinetas + 1
       Next I
       MsgBox "Foram incluídos " & qtMaquinetas & " processos de Mensalidade das Maquinetas no sistema de Contas, na data de " & CDate("01/" & TXT_MES & "/" & TXT_ANO) & ".", vbInformation, "Exportação de Mensalidade das Maquinetas"
       'FIM Mensalidade das Maquinetas
    End If
        
        
        MsgBox "Exportação concluída com sucesso!", vbInformation, "Exportação concluída"
    
sair:
    Exit Sub
err1:
    
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    
    Resume sair
End Sub

Private Sub cmdFaltas_Click()
    frm_Faltas.Show 1
End Sub

Private Sub cmdFiltrar_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
      Case 115: mnuAcessoTotal_Click 'F4
     End Select
End Sub

Private Sub cmdFixos_Click()
On Error GoTo err1
    '*** Gerar Lançamentos Fixos

        Dim adoFuncs As ADODB.Recordset
        Dim adoFixos As ADODB.Recordset
        
        Set adoFuncs = de.cnc.Execute("SELECT DISTINCT(TAB_FICHA_MENS.M_NFICHA) as FICHA, TAB_FICHA_MENS.M_F_COD as COD FROM TAB_FICHA_MENS, TAB_DESC_CALC_FIXO WHERE TAB_FICHA_MENS.M_F_COD = TAB_DESC_CALC_FIXO.CF_EMP_COD AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ")").Clone
    
        Do While Not adoFuncs.EOF
            Set adoFixos = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_EMP_COD = " & adoFuncs.Fields("COD")).Clone
        
            Do While Not adoFixos.EOF
                de.cnc.Execute ("DELETE FROM TAB_DESC_CALC Where (C_TP_CONTA = " & adoFixos.Fields("CF_TP_CONTA") & ") And (C_N_FICHA = " & adoFuncs.Fields("FICHA") & ") AND (C_NCRED = " & adoFixos.Fields("CF_CODIGO") & ")")
                de.cmdIncluirDescCalc2 Date, adoFuncs.Fields("FICHA"), adoFixos.Fields("CF_TP_CONTA"), adoFixos.Fields("CF_TP_OP"), adoFixos.Fields("CF_VALOR"), adoFixos.Fields("CF_DESC"), "0", adoFixos.Fields("CF_CODIGO"), "0", "0", adoFixos.Fields("CF_EMP_COD"), 0
                adoFixos.MoveNext
            Loop
          
            adoFuncs.MoveNext
        Loop

    ADOREG.Recordset.MoveFirst
        
    'Dados Contas
    Lancamentos
    
sair:
    Exit Sub
err1:
    
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    
    Resume sair
End Sub

Private Sub cmdFixosSaldos_Click()
Dim mes, ano As String
    
If de.rscmdSqlFixosSaldos_Grouping.State = 1 Then de.rscmdSqlFixosSaldos_Grouping.Close
    
    'frm_ESCOLHA_DATA.Show 1
    
    'dtIni = CDate(frm_ESCOLHA_DATA.TXT_DT_INICIAL)
    'dtFim = CDate(frm_ESCOLHA_DATA.TXT_DT_FINAL)
    
     ano = InputBox("Entre com o Ano:")
     mes = InputBox("Entre com o Mês:")

    If ano <> "" And mes <> "" Then
        de.cmdSqlFixosSaldos_Grouping ano, mes
    
        rptFixosSaldos.Sections("secCab").Controls("lbData").Caption = "(" & mes & "/" & ano & ")"
    
        'rptSalarioGerentes.Sections("SecCab").Controls("lbTitulo").Caption = "SAL. G (" & Month(dtIni) & ")"
 
        rptFixosSaldos.Show
    End If
End Sub

Private Sub cmdComissGerente_Click()
On Error Resume Next
    Dim vrVenda, vrFixo, vrMinimo, percComis, vrSalario, vrComis
    Dim vrVendaAnt, percVenda, vrBonus, vrSalarioBonus
    Dim vrMeta, vrMetaBonus, descMeta, wFixoMeta
    
    If Not isMesValido(txt_F_COD, TXT_MES, TXT_ANO) Then 'Verifica se é mês atual ou passado
        If MsgBox("Você está alterando uma ficha que NÃO É DO MÊS ATUAL. Deseja continuar mesmo assim?", vbYesNo, "Alteração de fichas") = vbNo Then
            Exit Sub
        End If
        If ADOREG.Recordset.Fields("M_BLOQ") Then
            MsgBox "Esta ficha não é do mês atual e está BLOQUEADA!", vbCritical
            Exit Sub
        End If
    End If
    
    If MsgBox("Deseja (re)gerar a comissão de todos os Gerentes do mês " & TXT_MES & "?", vbYesNo, "Gerar comissão") = vbNo Then
        Exit Sub
    End If
    
    
    cbMostrar.text = "GERENTE"
    cmdMostrar_Click
    
    ADOREG.Recordset.MoveFirst
    Do While Not ADOREG.Recordset.EOF
    
        'If IsNull(adoReg.Recordset.Fields("M_DT_DEM")) And adoReg.Recordset.Fields("M_LOGO") = "FH" Then
        'Or (adoReg.Recordset.Fields("M_LOGO") = "SV"))) Then
        If IsNull(ADOREG.Recordset.Fields("M_DT_DEM")) Then
    
            'COD da loja do gerente do ANO anterior
            Set w_ado_vendaAnt = de.cnc.Execute("SELECT TAB_VENDA.V_VR From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND TAB_FUNCIONARIO.F_DT_DEM IS NULL AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & Format(ADOREG.Recordset.Fields("M_MES"), "00") & " AND Right(TAB_VENDA.V_DATA,4) = " & CInt(ADOREG.Recordset.Fields("M_ANO")) - 1 & " AND TAB_FUNCIONARIO.F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD")).Clone
            If Not w_ado_vendaAnt.EOF Then
                vrVendaAnt = w_ado_vendaAnt.Fields(0)
            Else
                MsgBox "Não há lançamentos do logo " & W_ADO_FICHA.Fields("M_LOGO") & " para o período: " & Format(ADOREG.Recordset.Fields("M_MES"), "00") & " / " & CInt(ADOREG.Recordset.Fields("M_ANO")) - 1 & "! Ignorando...", vbCritical
            End If
            
            'COD da loja do cx do ANO atual
            Set w_ado_venda = de.cnc.Execute("SELECT TAB_VENDA.V_VR From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND TAB_FUNCIONARIO.F_DT_DEM IS NULL AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & Format(ADOREG.Recordset.Fields("M_MES"), "00") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD")).Clone
            If Not w_ado_venda.EOF Then
                vrVenda = w_ado_venda.Fields(0)
            Else
                vrVenda = 0
                MsgBox "Não há lançamentos do logo " & W_ADO_FICHA.Fields("M_LOGO") & " para o período: " & Format(ADOREG.Recordset.Fields("M_MES"), "00") & " / " & ADOREG.Recordset.Fields("M_ANO") & "! Ignorando...", vbCritical
            End If
        
            If Not w_ado_vendaAnt.EOF And Not w_ado_venda.EOF Then
       
            'vrVenda = de.cnc.Execute("SELECT TAB_VENDA.V_VR From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND TAB_FUNCIONARIO.F_DEM_OK = 0 AND TAB_FUNCIONARIO.F_L AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & adoReg.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & adoReg.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & adoReg.Recordset.Fields("M_F_COD")).Fields(0)
            'vrVenda = de.cnc.Execute("SELECT TAB_VENDA.V_VR From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND TAB_FUNCIONARIO.F_DEM_OK = 0 AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & Format(adoReg.Recordset.Fields("M_MES"), "00") & " AND Right(TAB_VENDA.V_DATA,4) = " & adoReg.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & adoReg.Recordset.Fields("M_F_COD")).Fields(0)
        
            'If vrVenda <> "" Then
                vrFixo = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_VR_FIXO From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & ADOREG.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD")).Fields(0)
                vrMinimo = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_VR_MINIMO From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & ADOREG.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD")).Fields(0)
                percComis = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_COMIS From TAB_FUNCIONARIO, TAB_VENDA WHERE TAB_FUNCIONARIO.F_LOJA = TAB_VENDA.V_F_LOJA AND Right(Left(TAB_VENDA.V_DATA,5),2) = " & ADOREG.Recordset.Fields("M_MES") & " AND Right(TAB_VENDA.V_DATA,4) = " & ADOREG.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD")).Fields(0)
                
                
                 '**** BONUS DE META
                 vrMeta = 999
                 'vrMeta = de.cnc.Execute("SELECT MT_VALOR FROM TAB_META INNER JOIN TAB_FUNCIONARIO ON MT_F_LOJA = TAB_FUNCIONARIO.F_LOJA WHERE MT_MES=" & adoReg.Recordset.Fields("M_MES") & " AND MT_ANO=" & adoReg.Recordset.Fields("M_ANO") & " AND TAB_FUNCIONARIO.F_Codigo=" & W_ADO_FICHA.Fields("M_F_COD")).Fields(0)
                 If ADOREG.Recordset.Fields("M_LOGO") <> "XX" Then
                     vrMeta = de.cnc.Execute("SELECT TAB_META.MT_VALOR FROM TAB_META, Lojb010 INNER JOIN TAB_FUNCIONARIO ON Lojb010.COD_LOJ = TAB_FUNCIONARIO.F_Cod_L WHERE (((TAB_META.MT_MES)=" & ADOREG.Recordset.Fields("M_MES") & ") AND ((TAB_META.MT_ANO)=" & ADOREG.Recordset.Fields("M_ANO") & ") AND ((TAB_FUNCIONARIO.F_Codigo)=" & ADOREG.Recordset.Fields("M_F_COD") & ") AND ((Lojb010.NUM)=Right([tab_meta].[mt_f_loja],2)))").Fields(0)
                 End If
                 '****
                
                
                
                vrSalario = vrFixo + ((vrVenda * 1000) * (percComis / 100))
                vrComis = (vrVenda * 1000) * (percComis / 100)
                'percVenda = FormatNumber((100 - (vrVenda / vrVendaAnt * 100)) * -1, 2)
                percVenda = FormatNumber((100 - (vrVenda / vrMeta * 100)) * -1, 2)
                
                'If percVenda > 0 Then
                    'Lojas em promoção ganham metade do bonus
                    'If ADOREG.Recordset.Fields("M_LOGO") = "AC" Or ADOREG.Recordset.Fields("M_LOGO") = "IJ" Or ADOREG.Recordset.Fields("M_LOGO") = "FH" Then
                    '    percVenda = percVenda / 2
                    'End If
    
                    'If percVenda >= 40 Then percVenda = percVenda / 2
                    'If adoReg.Recordset.Fields("M_MES") = "12" Then
                        'vrBonus = percVenda * 10
                        'vrBonus = 0
                    'Else
                        'If (vrVenda < 40) Then
                            'vrBonus = percVenda * 25
                        'Else
                            'vrBonus = percVenda * 50
                        'End If
                    'End If
                    
                    'Algumas gerentes em algumas lojas não ganham bonus
                    'If (adoReg.Recordset.Fields("M_LOGO") = "LP") Then
                    '        vrBonus = 0
                    'End If
                    
                'Else
                    vrBonus = 0
                'End If
                
                'If vrBonus > 1000 And adoReg.Recordset.Fields("M_MES") <> "12" Then vrBonus = 1000
                'If percVenda >= 5 Then
                    'If vrBonus >= 250 Then
                        'vrBonus = 250
                    'Else
                        'vrBonus = 125
                    'End If
               ' End If
                
                
                vrSalarioBonus = vrSalario
                'vrSalarioBonus = vrSalario + vrBonus
                
              ' If vrSalario < vrMinimo Then
              '      'vrSalarioBonus = vrMinimo + vrBonus
              '      vrSalarioBonus = vrMinimo
              '  End If
                
                 '**** BONUS DE META
             vrMetaBonus = 0
             wFixoMeta = 0
             wPercMeta = 0
                 
              If ADOREG.Recordset.Fields("m_mes") <= 12 Then
                   
                If vrVenda >= (vrMeta) Then
                    wFixoMeta = vrSalario * 0.1
                    If (((vrVenda / vrMeta) - 1) * 100) >= 5 Then
                        vrMetaBonus = vrSalario * 0.05
                        wPercMeta = 5
                    Else
                        vrMetaBonus = vrSalario * ((vrVenda / vrMeta) - 1)
                        wPercMeta = ((vrVenda / vrMeta) - 1) * 100
                    End If
                    descMeta = " + Bonus Meta(" & Format(vrSalario, "0.00") & " * " & Format(wPercMeta, "0.0") & "% = " & Format(vrMetaBonus, "0.00") & " + Fixo Meta(" & Format(wFixoMeta, "0.00") & ") = " & Format(vrMetaBonus + wFixoMeta, "0.00") & ") "
                Else
                   descMeta = ""
                End If
                
                If vrSalario < vrMinimo Then
                    vrSalarioBonus = vrMinimo + vrMetaBonus + wFixoMeta
                Else
                    vrSalarioBonus = vrSalarioBonus + vrMetaBonus + wFixoMeta
                End If

              End If
                
                If vrSalario < vrMinimo Then
                    wDesc = "**NÃO ATINGIU O MÍNIMO** (" & Format(vrVenda, "0.00") & " x " & percComis & "% = " & Format(vrComis, "0.00") & ") + " & Format(vrFixo, "0.00") & " = " & Format(vrSalario, "0.00") & " < Mínimo(" & Format(vrMinimo, "0.00") & ") " & descMeta & " = " & Format(vrSalarioBonus, "0.00")
                Else
                    wDesc = "(" & Format(vrVenda, "0.00") & " x " & percComis & "% = " & Format(vrComis, "0.00") & ") + " & Format(vrFixo, "0.00") & " = " & Format(vrSalario, "0.00") & descMeta & " = " & Format(vrSalarioBonus, "0.00")
                End If
                
                de.cnc.Execute ("DELETE FROM TAB_DESC_CALC Where (C_TP_CONTA = 20) And (C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFicha") & ")")
                de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFicha"), 20, "+", vrSalarioBonus, wDesc, "", "0", "0", "0", "0"
                
            End If
        End If
        ADOREG.Recordset.MoveNext
    Loop
    
    ADOREG.Recordset.MoveFirst
    'Dados Contas
    Lancamentos
    
End Sub

Private Sub cmdGerarSalarioTodos2_Click()

End Sub

Private Sub cmdFixoX_Click()
On Error GoTo err1
    '*** Gerar Lançamentos Fixos

        Dim adoFuncs As ADODB.Recordset
        Dim adoFixos As ADODB.Recordset
        
        Set adoFuncs = de.cnc.Execute("SELECT DISTINCT(TAB_FICHA_MENS.M_NFICHA) as FICHA, TAB_FICHA_MENS.M_F_COD as COD FROM TAB_FICHA_MENS, TAB_DESC_CALC_FIXO WHERE TAB_FICHA_MENS.M_F_COD = TAB_DESC_CALC_FIXO.CF_EMP_COD AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_PG_VT = True) ").Clone
    
        Do While Not adoFuncs.EOF

            Dim fichaAtual As String
            Dim ultimoFixo As String

            fichaAtual = adoFuncs.Fields("FICHA")
            
            de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & fichaAtual & " AND (C_TP_CONTA = 111)")
            de.cnc.Execute ("DELETE FROM TAB_DESC_CALC_FIXO WHERE CF_EMP_COD = " & adoFuncs.Fields("COD") & " AND (CF_TP_CONTA = 111)")
            
            de.cmdIncluirDescCalcFixo Now(), adoFuncs.Fields("COD"), "111", "=", "0", "Pagto. de passes (vale transporte) [GERADO AUTOMATICAMENTE]"
            ultimoFixo = de.cnc.Execute("SELECT Max([CF_CODIGO]) FROM TAB_DESC_CALC_FIXO").Fields(0)
            Set adoFixos = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_CODIGO = " & ultimoFixo).Clone
            de.cmdIncluirDescCalc2 Date, fichaAtual, adoFixos.Fields("CF_TP_CONTA"), adoFixos.Fields("CF_TP_OP"), adoFixos.Fields("CF_VALOR"), adoFixos.Fields("CF_DESC"), "0", adoFixos.Fields("CF_CODIGO"), "0", "0", adoFixos.Fields("CF_EMP_COD"), 0

            adoFuncs.MoveNext
        Loop

    ADOREG.Recordset.MoveFirst

    'Dados Contas
    Lancamentos
 
sair:
    Exit Sub
err1:
    
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    
    Resume sair
End Sub

Private Sub cmdMostrar_Click()
    Select Case cbMostrar.text
        Case "TODOS":
            Op_Click 5
        Case "VENDEDOR":
            txt_Pesq = "'V'"
            FILTRAR 8
        Case "CAIXA":
            txt_Pesq = "'C'"
            FILTRAR 8
        Case "2º CAIXA":
            txt_Pesq = "'2'"
            FILTRAR 8
        Case "CX EXTRA":
            txt_Pesq = "'X'"
            FILTRAR 8
        Case "SEGURANÇA":
            txt_Pesq = "'R'"
            FILTRAR 8
        Case "GERENTE":
            txt_Pesq = "'G'"
            FILTRAR 8
        Case "GER RODA":
            txt_Pesq = "'D'"
            FILTRAR 8
        Case "SUPERVISOR":
            txt_Pesq = "'S'"
            FILTRAR 8
        Case "RP":
            txt_Pesq = "'O'"
            FILTRAR 8
        Case "(NÃO BLOQ.)":
            txt_Pesq = "'###'"
            FILTRAR 10
    End Select
    
    txt_Pesq = ""
    Lancamentos
End Sub

Private Sub cmdRelAdmin_Click()
    
    cmdComisMwts.Visible = Not cmdComisMwts.Visible
    cmdComisCx.Visible = Not cmdComisCx.Visible
    cmdComixCxEXT.Visible = Not cmdComixCxEXT.Visible
    cmd13.Visible = Not cmd13.Visible
    cmdSaldo.Visible = Not cmdSaldo.Visible
    cmdComissGerente.Visible = Not cmdComissGerente.Visible
    cmdAddLanç_SalFTodos.Visible = Not cmdAddLanç_SalFTodos.Visible
    cmdEmprestimo.Visible = Not cmdEmprestimo.Visible
    cmdFixos.Visible = Not cmdFixos.Visible
    cmdExportarContas.Visible = Not cmdExportarContas.Visible
    cmdDespesa.Visible = Not cmdDespesa.Visible
    cmdDespesa99.Visible = Not cmdDespesa99.Visible
    cmdBonus.Visible = Not cmdBonus.Visible
    cmdComerciario.Visible = Not cmdComerciario.Visible
    cmdSindical.Visible = Not cmdSindical.Visible
    cmdFixoX.Visible = Not cmdFixoX.Visible
    cmdDtDEM.Visible = Not cmdDtDEM.Visible
    cmdDelFichas.Visible = Not cmdDelFichas.Visible

End Sub

Private Sub cmdRelQtdeTipo_Click()
On Error GoTo err1

    If de.rscmdSqlQtdeTipo.State = 1 Then de.rscmdSqlQtdeTipo.Close
    de.cmdSqlQtdeTipo TXT_MES, TXT_ANO
    
    rptQtdeTipo.Sections("SecCab").Controls("lbTitulo").Caption = "FUNCIONÁRIOS por FUNÇÃO (" & TXT_MES & "/" & TXT_ANO & ")"
    'rptQtdeTipo.Sections("SecCab").Controls("lbData").Caption = Format(Date, "DD=MM") & " " & Format(Time, "hh=mm")
    
    rptQtdeTipo.Show
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub cmdSalarioCX_Click()
    If de.rscmdSqlSalarioCxNOVO.State = 1 Then de.rscmdSqlSalarioCxNOVO.Close
     
    
    de.cmdSqlSalarioCxNOVO TXT_ANO, TXT_MES
    rptSalarioCxNOVO.Sections("SecCab").Controls("lbTitulo").Caption = "SAL. CXs. (" & TXT_MES & ")"
    
         
    rptSalarioCxNOVO.Show
End Sub

Private Sub cmdSalarioGerente_Click()
Dim dtIni, dtFim As Date
    
'If de.rscmdSqlSalarioGerentes.State = 1 Then de.rscmdSqlSalarioGerentes.Close
    
    frm_ESCOLHA_DATA.Show 1
    
    dtIni = CDate(frm_ESCOLHA_DATA.TXT_DT_INICIAL)
    dtFim = CDate(frm_ESCOLHA_DATA.TXT_DT_FINAL)
    
     'dtIni = InputBox("Entre com a Data Inicial:", , Format(Date, "DD/MM/YYYY"))
     'dtFim = InputBox("Entre com a Data Final:", , Format(Date, "DD/MM/YYYY"))
     
    If de.rscmdSqlSalarioGerentes.State = 1 Then de.rscmdSqlSalarioGerentes.Close
     
    de.cmdSqlSalarioGerentes dtIni, dtFim
    rptSalarioGerentes.Sections("SecCab").Controls("lbTitulo").Caption = "SAL. G (" & Month(dtIni) & ")"
    'rptSalarioGerentes.Sections("SecCab").Controls("lbData").Caption = Format(Date, "DD=MM") & " " & Format(Time, "hh=mm")
         
    rptSalarioGerentes.Show
    
End Sub



Private Sub Command1_Click()
    ADOREG.Recordset.MoveFirst
    ADOREG.Recordset.Find "m_f_cod = " & InputBox("COD FUNC:"), , adSearchForward
End Sub

Private Sub ctr_Button1_Click()

End Sub

Private Sub ctr_Button2_Click()

End Sub



Private Sub cmdSaldo_Click()
Dim vrVenda, vrFixo, vrMinimo, percComis, vrSalario, vrComis
Dim ww_mes, ww_ano, qtdeSaldoAdded
On Error Resume Next

    
    If Not isMesValido(txt_F_COD, ADOREG.Recordset.Fields("M_MES"), ADOREG.Recordset.Fields("M_ANO")) Then 'Verifica se é mês atual ou passado
        If MsgBox("Você está alterando uma ficha que NÃO É DO MÊS ATUAL. Deseja continuar mesmo assim?", vbYesNo, "Alteração de fichas") = vbNo Then
            Exit Sub
        End If
        'If adoReg.Recordset.Fields("M_BLOQ") Then
            'MsgBox "Esta ficha não é do mês atual e está BLOQUEADA!", vbCritical
            'Exit Sub
        'End If
    End If
    
    
    If MsgBox("Deseja regerar o Saldo Devedor de todos do mês " & ADOREG.Recordset.Fields("M_MES") & "?", vbYesNo, "Gerar saldo") = vbNo Then
        Exit Sub
    End If
    
    qtdeSaldoAdded = 0
    
    'Voltando para o mês anterior
    bt_VoltarDT_Click


    ww_mes = ADOREG.Recordset.Fields("M_MES") + 1
    If ww_mes = 13 Then
        ww_mes = 1
        ww_ano = ADOREG.Recordset.Fields("M_ANO") + 1
    Else
        ww_ano = ADOREG.Recordset.Fields("M_ANO")
    End If
    
    ADOREG.Recordset.MoveFirst
    Do While Not ADOREG.Recordset.EOF
    Dim ADO_TOTAL As ADODB.Recordset
    Dim wTXT_MAIS
    Dim wTXT_MENOS
    Dim wTXT_TOTAL
      
      wTXT_MAIS = 0
        wTXT_MENOS = 0
        wTXT_TOTAL = 0
        
        Set ADO_TOTAL = ADO_LANC.Recordset.Clone
        
        If Not ADO_TOTAL.EOF Then
            ADO_TOTAL.MoveFirst
            Do While Not ADO_TOTAL.EOF
                If ADO_TOTAL.Fields("VALOR") >= 0 And ADO_TOTAL.Fields("OP") = "+" Then
                    wTXT_MAIS = CDbl(wTXT_MAIS) + ADO_TOTAL.Fields("VALOR")
                ElseIf ADO_TOTAL.Fields("VALOR") < 0 And ADO_TOTAL.Fields("OP") = "-" Then
                    wTXT_MENOS = CDbl(wTXT_MENOS) + ADO_TOTAL.Fields("VALOR")
                End If
                ADO_TOTAL.MoveNext
            Loop
            
            wTXT_TOTAL = CDbl(wTXT_MAIS) + CDbl(wTXT_MENOS)
        End If
 
        Dim proxFicha
        Dim w_desc
 
        proxFicha = de.cnc.Execute("SELECT M_NFICHA From TAB_FICHA_MENS WHERE M_ANO = " & ww_ano & " AND M_MES = " & ww_mes & " AND M_F_COD = " & ADOREG.Recordset.Fields("M_F_COD")).Fields(0)
        de.cnc.Execute ("DELETE FROM TAB_DESC_CALC Where (C_TP_CONTA = 14) And (C_N_FICHA = " & proxFicha & ")")
        'If wTXT_TOTAL < 0 And ADOREG.Recordset.Fields("M_LOGO") <> "RP" And (IsNull(ADOREG.Recordset.Fields("M_DT_ACF")) Or ADOREG.Recordset.Fields("M_DT_ACF") = "") Then
        If wTXT_TOTAL < 0 And ADOREG.Recordset.Fields("M_LOGO") <> "RP" And Not (IsEmpty(proxFicha)) Then
            w_desc = "Pg. Saldo Dev.: " & Format(wTXT_TOTAL, "R$ 0.00")
            de.cmdIncluirDescCalcVistado Date, proxFicha, 14, "-", wTXT_TOTAL, w_desc, "", "0", "0", "0", ADOREG.Recordset.Fields("M_F_COD")
            qtdeSaldoAdded = qtdeSaldoAdded + 1
        End If

        ADOREG.Recordset.MoveNext
    Loop
    
    'Retornando para o mês atual
    bt_AvaDT_Click
    
    MsgBox "No mês " & ww_mes & " houveram " & qtdeSaldoAdded & " fichas com lançamento de Saldo Negativo!", vbInformation, "Saldo Negativo"
    
sair:
    Exit Sub
err1:
    If Not Err.Number = 3705 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub cmdteste_Click()

End Sub

Private Sub cmdSindical_Click()
On Error GoTo err1
    Dim wValorSindical
    Dim wDesc
    
    ADOREG.Recordset.MoveFirst
    Do While Not ADOREG.Recordset.EOF
        
        If ADOREG.Recordset.Fields("M_LOGO").value <> "RP" Then
            
            wValorSindical = ADOREG.Recordset.Fields("f_vpiso").value
            
            If wValorSindical > 0 Then
                wValorSindical = Format((wValorSindical / 30) * -1, "0.00") 'Calcula Contribuição Sindical
                wDesc = Format(ADOREG.Recordset.Fields("f_vpiso").value, "0.00") & " / 30 =  " & Format(wValorSindical, "0.00")
                de.cnc.Execute ("DELETE FROM TAB_DESC_CALC Where (C_TP_CONTA = 70) And (C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFicha") & ")")
                de.cmdIncluirDescCalc Date, ADOREG.Recordset.Fields("M_NFicha"), 70, "-", wValorSindical, wDesc, "", "0", "0", "0", "0"
            End If
            
        End If
        
        wValorSindical = 0
        wDesc = ""

        ADOREG.Recordset.MoveNext
    Loop
    
    ADOREG.Recordset.MoveFirst
        
    'Dados Contas
    Lancamentos
    
    
sair:
    Exit Sub
err1:
    
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    
    Resume sair
End Sub

Private Sub flexGRID_L_DblClick()
    If flexGRID_L.RowSel <> 0 Then CONTA
End Sub

Private Sub flexGRID_L_KeyDown(KeyCode As Integer, Shift As Integer)
    If flexGRID_L.RowSel <> 0 Then
      If (UCase(frmLogin.txtUserName) = UCase(NomeMestre) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre2) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre3)) And Shift = 0 And KeyCode <> 13 Then
            'F7
           Select Case KeyCode
            Case 115: mnuAcessoTotal_Click 'F4
            Case 118: mnuVis_Click  'F7
            Case 119: mnuRem_Click  'F8
            Case 122: mnuVist_Click 'F11
            Case 123: mnuRemT_Click 'F12
          End Select
        ElseIf Shift <> 2 And KeyCode = 13 Then
            If Grid.Enabled = True Then
                Grid.SetFocus
            Else
                txt_DT_ADM.SetFocus
            End If
        End If
    End If
End Sub



Private Sub flexGRID_L_KeyPress(KeyAscii As Integer)
    If flexGRID_L.RowSel <> 0 Then
        Select Case KeyAscii
        ' Editar ao teclar ENTER
        Case vbKeyReturn
            If ADOREG.Recordset.Fields("M_BLOQ") = False Then
            
                If Not isMesValido(txt_F_COD, TXT_MES, TXT_ANO) Then 'Verifica se é mês atual ou passado
                    If MsgBox("Você está alterando uma ficha que NÃO É DO MÊS ATUAL. Deseja continuar mesmo assim?", vbYesNo, "Alteração de ficha") = vbNo Then
                        Exit Sub
                    End If
                    If ADOREG.Recordset.Fields("M_BLOQ") Then
                        MsgBox "Esta ficha não é do mês atual e está BLOQUEADA!", vbCritical
                        Exit Sub
                    End If
                End If
                
                If flexGRID_L.TextMatrix(flexGRID_L.RowSel, 6) <> "Não" Then
                    If ADOREG.Recordset.Fields("M_BLOQ") Then
                        MsgBox "Esta ficha não é do mês atual e está BLOQUEADA!", vbCritical
                        Exit Sub
                    End If
                    frm_Habilitar.Show 1
                    w_PSS = frm_Habilitar.txt_Pss
                    Else
                        w_PSS = w_PassWordLib
                    End If
                    
                    If w_PSS = w_PassWordLib Then
                        KeyAscii = 0
                        ExibirCelula
                    Else
                        MsgBox "Senha de Liberação Incorreta!", vbCritical
                    End If
            Else
                MsgBox "Não é possível alterar uma ficha anterior ao mês passado!", vbExclamation
            End If
        End Select
    End If
        
End Sub

Private Sub flexGRID_L_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If flexGRID_L.RowSel <> 0 Then
        If Button = 2 And (UCase(frmLogin.txtUserName) = UCase(NomeMestre) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre2) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre3)) And CK_ACF = 0 Then
           If Not isMesValido(txt_F_COD, TXT_MES, TXT_ANO) Then 'Verifica se é mês atual ou passado
                If MsgBox("Você está alterando uma ficha que NÃO É DO MÊS ATUAL. Deseja continuar mesmo assim?", vbYesNo, "Alteração de ficha") = vbNo Then
                    Exit Sub
                End If
                If ADOREG.Recordset.Fields("M_BLOQ") Then
                    MsgBox "Esta ficha não é do mês atual e está BLOQUEADA!", vbCritical
                    Exit Sub
                End If
            End If
            PopupMenu mnu
        End If
    End If
End Sub



Private Sub Form_Activate()
On Error GoTo err1
    Me.WindowState = 2
    
w_reset_tipo = True

flexGRID_L.ColWidth(0) = 880 'data
flexGRID_L.ColWidth(1) = 420 'tp_conta
flexGRID_L.ColWidth(2) = 3450 'tp_desc (descrição do tipo da conta)
flexGRID_L.ColWidth(3) = 4000 'conta
flexGRID_L.ColWidth(4) = 1080 'valor
flexGRID_L.ColWidth(5) = 330 'op
flexGRID_L.ColWidth(6) = 550 'visto
flexGRID_L.ColWidth(7) = 0 'codigo lancamento
flexGRID_L.ColWidth(8) = 0 'codigo ncred (codigo do fixo)

BarraF.Buttons("desbloquear").Enabled = False
cmdDesbloquear.Visible = False
'BarraF.Buttons("gerente").Enabled = False

'If (adoReg.Recordset.Fields("M_TIPO") = "V" Or adoReg.Recordset.Fields("M_TIPO") = "C") Or acessoTotal() Then
'    lbl_SaldoAnt.Visible = True
'    txt_SaldoAnt.Visible = True
'
'    TXT_TOTAL.Visible = True
'    lbl_total.Visible = True
'
'Else
'    lbl_SaldoAnt.Visible = False
'    txt_SaldoAnt.Visible = False
'
'    TXT_TOTAL.Visible = False
'    lbl_total.Visible = False
'End If
 
 If acessoTotal() Then
     
'    BarraF.Buttons("gerente").Enabled = True
    lblNotas.Visible = True
    txt_notas.Visible = True
 
    cmdSalarioGerente.Visible = True
    'cmdFixosSaldos.Visible = True
    
    txtQtdeLimiteV.Enabled = True
    
    If (UCase(w_usuario) = UCase(NomeMestre) Or UCase(w_usuario) = UCase(NomeMestre2) Or UCase(w_usuario) = UCase(NomeMestre3)) Then
        BarraF.Buttons("desbloquear").Enabled = True
        cmdDesbloquear.Visible = True
        'cmdRelAdmin.Visible = True
    End If
    
    'If UCase(p_Usuario) = "BEL" Then
    '    cmdRelAdmin.Visible = True
    'End If
    

    cmdRelAdmin.Visible = True
    cmdCopiarFicha.Visible = True
    
    'If UCase(p_Usuario) = "RODRIGO" Or UCase(p_Usuario) = "KELEN" Then
        mnuAcessoTotal_Click
    'End If

    
    
End If
    

    Editar 0
    Cancelar
    
    If BarraF.Buttons("salvar").Enabled = False And BarraF.Buttons("editar").Enabled = False Then BarraF.Buttons("editar").Enabled = True

        

sair:
   ' Salvar
    Exit Sub
err1:
    If Err.Number <> 3705 And Err.Number <> -2147217864 Then
        'MsgBox Err.Number & " : " & Err.Description, vbCritical
    End If
    
    'Cancelar
    Cancelar

    Resume sair
    
    
End Sub

Private Sub Form_Activate2()
On Error Resume Next
      'Dados Contas
      
      
      
    Lancamentos
    
On Error GoTo err1
    
    
    'Cancelar
    'Cancelar
    'Refresh_dados

    
    If BarraF.Buttons("salvar").Enabled = False And BarraF.Buttons("editar").Enabled = False Then BarraF.Buttons("editar").Enabled = True
    
    'If UCase(frmLogin.txtUserName) = UCase(NomeMestre) Then BarraF.Buttons("dupla").Visible = True
   
    '*** CALCULA O TOTAL - APÓS O NOVO VALOR ***
    W_MAIS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MAIS FROM TAB_DESC_CALC  WHERE C_Tp_Op = '+' and C_VALOR > 0 AND (C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.ADOREG.Recordset.Fields("M_NFICHA") & ")").Fields("MAIS")
    W_MENOS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MENOS FROM TAB_DESC_CALC WHERE C_Tp_Op = '-' and C_VALOR < 0 AND (C_N_FICHA = " & frm_Alt_Fic_Mensal_VIS.ADOREG.Recordset.Fields("M_NFICHA") & ")").Fields("MENOS")
    
    W_TOTAL = IIf(IsNull(W_MENOS), 0, W_MENOS) + IIf(IsNull(W_MAIS), 0, W_MAIS)
 
    If ADOREG.Recordset.Fields("m_TOTAL") <> CDbl(W_TOTAL) Then
        TXT_TOTAL = Format(W_TOTAL, "R$ 0.00")
'        adoReg.Recordset.Fields("m_TOTAL") = TXT_TOTAL
        ADOREG.Recordset.UpdateBatch 'adAffectCurrent
    End If

    
    If ADOREG.Recordset.Fields("m_TOTAL") < 0 Then
        TXT_TOTAL.ForeColor = vbRed
    Else
        TXT_TOTAL.ForeColor = vbWhite
    End If

    'TXT_FUNC.SetFocus
    'w_F5 = False
   

    W_CK_DEM = True
    
sair:
   ' Salvar
    Exit Sub
err1:
    If Err.Number <> 3705 And Err.Number <> -2147217864 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
'    Set ADO_LANC.Recordset = de.cnc.Execute("SELECT TAB_DESC_CALC.C_DT AS DATA, 'CT: ' + TAB_TP_CONTA.TP_DESC + '     DESC: ' + TAB_DESC_CALC.C_DESC AS CONTA, TAB_DESC_CALC.C_VALOR AS VALOR, TAB_DESC_CALC.C_TP_OP AS OP FROM TAB_DESC_CALC, TAB_TP_CONTA WHERE TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_DESC_CALC.C_N_FICHA = " & frm_Alt_Fic_Mensal_Visualizar.ADOREG.Recordset.Fields("M_NFICHA") & ")").Clone
    Cancelar
    Cancelar

    Resume sair
End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
      Case 115: mnuAcessoTotal_Click 'F4
     End Select
End Sub

Sub Form_Load()
On Error GoTo err1

    If w_umaVez = 0 Then
        frmSplash.PB.value = 25
        frmSplash.PB.text = "Carregando " & frmSplash.PB.value & "%"
    End If

    w_SN_Total = True
    
    txt_PMes = Format(Date, "MM")
    If txt_PMes = 0 Then
        txt_PMes = 12
        txt_PAno = Format(Date, "YYYY") - 1
    Else
        txt_PAno = Format(Date, "YYYY")
    End If
    
    
   On Error Resume Next
    'Set adoReg.Recordset = Nothing
    'Set ADO_LANC.Recordset = Nothing
    'Set adoReg.Recordset = Nothing
    'Set ADO_LANC.Recordset = Nothing
    If de.rsTAB_FICHA_MENS.State = 1 Then de.rsTAB_FICHA_MENS.Close
    de.TAB_FICHA_MENS
    If w_umaVez = 0 Then
        frmSplash.PB.value = 45
        frmSplash.PB.text = "Carregando " & frmSplash.PB.value & "%"
    End If
    
    If de.rsTAB_DESC_CALC.State = 1 Then de.rsTAB_DESC_CALC.Close
    de.TAB_DESC_CALC
    If w_umaVez = 0 Then
        frmSplash.PB.value = 65
        frmSplash.PB.text = "Carregando " & frmSplash.PB.value & "%"
    End If
    
    If de.rscmdSqlVisAltFichas.State = 1 Then de.rscmdSqlVisAltFichas.Close
    de.rscmdSqlVisAltFichas.Resync
    
    If w_umaVez = 0 Then
        frmSplash.PB.value = 75
        frmSplash.PB.text = "Carregando " & frmSplash.PB.value & "%"
    End If
    
On Error GoTo err1
          
    
    If de.rscmdSqlVisAltFichas.State = 1 Then de.rscmdSqlVisAltFichas.Close
    de.cmdSqlVisAltFichas txt_PMes, txt_PAno
    
    If w_umaVez = 0 Then
        frmSplash.PB.value = 85
        frmSplash.PB.text = "Carregando " & frmSplash.PB.value & "%"
    End If
    
    
    TXT_DATA = Date
    Do While de.rscmdSqlVisAltFichas.EOF
    
        TXT_DATA = CVDate("01/" & Format(TXT_DATA, "MM/YYYY")) - 1
        txt_PMes = Format(TXT_DATA, "MM")
        txt_PAno = Format(TXT_DATA, "YYYY")
    
        If de.rscmdSqlVisAltFichas.State = 1 Then de.rscmdSqlVisAltFichas.Close
        de.cmdSqlVisAltFichas txt_PMes, txt_PAno
        
    Loop
    
    Set ADOREG.Recordset = de.rscmdSqlVisAltFichas.Clone
    
    ADOREG.Recordset.MoveFirst
    
    If w_umaVez = 0 Then
        frmSplash.PB.value = 100
        frmSplash.PB.text = "Carregando " & frmSplash.PB.value & "%"
        w_umaVez = 1
    End If
    
    'Lancamentos
'    Set ADO_LANC.Recordset = ADOREG.Recordset.Fields("cmdSqlVisAltContas").UnderlyingValue
    
    
    V_MOVE = False

    If TXT_FILTRO <> "" Then ADOREG.Recordset.Filter = TXT_FILTRO
    
    V_MOVE = True
    
    'Saldo restante da ficha
    W_SALDO = de.cnc.Execute("Select F_SALDO_ANT FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
    txt_SaldoAnt = IIf(IsNull(W_SALDO), Format(0, "R$ 0.00"), Format(W_SALDO, "R$ 0.00"))
    'Saldo DO EMPRESTIMO
'    w_Emprest = de.cnc.Execute("Select F_EMPRESTIMO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
'    txt_SaldoEmp = IIf(IsNull(w_Emprest), Format(0, "R$ 0.00"), Format(w_Emprest, "R$ 0.00"))

    If txt_SaldoAnt < 0 Then
        txt_SaldoAnt.ForeColor = vbRed
    Else
        txt_SaldoAnt.ForeColor = vbBlue
    End If
        
    If de.rscmdBase.State = 1 Then ADOREG.Refresh
    'Editar

    '**** HABILITA P/ PL ***
    If (UCase(frmLogin.txtUserName) = UCase(NomeMestre) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre2) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre3)) Then
        Me.BarraF.Buttons("vistar").Visible = True
        CK_DEM.Enabled = True
    End If
    
    'Abre já com o filtro por B como padrão
    Op_Click 1
    
    Editar 0
    
sair:
    Exit Sub
err1:
    If Not Err.Number = 3705 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

'*** Caption no navegador ***
Private Sub adoReg_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

On Error Resume Next

            txt_VPiso = 0
            txt_VPiso = de.cnc.Execute("SELECT F_VPISO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
            txt_VPiso = Format(txt_VPiso, "0.00")
            txt_VPiso_R = 0
            txt_VPiso_R = de.cnc.Execute("SELECT F_VPISO_R FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
            txt_VPiso_R = Format(txt_VPiso_R, "0.00")
 

On Error GoTo err1

    'Hab e desab. o botão de add lanç. sal. fam.
    cmdAddLanç_SalF.Enabled = ADOREG.Recordset.Fields("M_PG_SAL_FAM")
    
    'If w_SN_Total = True And (Not adoReg.Recordset.EOF And Not adoReg.Recordset.BOF) And (Op(5).value = False And adReason <> 7) Then
    If w_SN_Total = True And (Not ADOREG.Recordset.EOF And Not ADOREG.Recordset.BOF) And (adReason <> 7) Then
        
        ADOREG.Caption = "REGISTRO : " & ADOREG.Recordset.AbsolutePosition & " / " & ADOREG.Recordset.RecordCount & IIf(W_LD_FILTRO = True, " (FILTRADO)", "")
            
            
            '*** DESABILITA O EDITAR ****
            
            If ADOREG.Recordset.Fields("M_BLOQ") = True And Not (UCase(frmLogin.txtUserName) = UCase(NomeMestre) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre2) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre3)) Then
                 BarraF.Buttons("editar").Enabled = False
            Else
                 BarraF.Buttons("editar").Enabled = True
            End If
            
            If ADOREG.Recordset.Fields("m_TOTAL") < 0 Then
                TXT_TOTAL.ForeColor = vbRed
            Else
                TXT_TOTAL.ForeColor = vbWhite
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
        
            'Saldo DO EMPRESTIMO
'            If de.rsTAB_FUNCIONARIO.State = 1 Then de.rsTAB_FUNCIONARIO.Requery
 '           w_Emprest = de.cnc.Execute("Select F_EMPRESTIMO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & adoReg.Recordset.Fields("M_F_COD") & "").Fields(0)
  '          txt_SaldoEmp = IIf(IsNull(w_Emprest), Format(0, "R$ 0.00"), Format(w_Emprest, "R$ 0.00"))
            
            'Dados Contas
            'If de.rscmdSqlVisAltContas.State = 1 Then de.rscmdSqlVisAltContas.Close
            'de.cmdSqlVisAltContas adoReg.Recordset.Fields("M_NFICHA")
            'Set ADO_LANC.Recordset = de.rscmdSqlVisAltContas.Clone

    
        End If
            
         If Not ADOREG.Recordset.EOF Then
            '***Carimbo***
            If (Len(ADOREG.Recordset.Fields("M_DT_ACF")) > 4) And (ADOREG.Recordset.Fields("M_BLOQ") = -1) Then
                CARIMBO.Visible = True
            Else
                CARIMBO.Visible = False
            End If
                    
            
            If ADOREG.Recordset.Fields("M_BLOQ") = True And IsNull(ADOREG.Recordset.Fields("M_DT_DEM")) Then
                frmBloq.Visible = True
                Shape1.BackColor = &HFFC0C0
            Else
                frmBloq.Visible = False
                Shape1.BackColor = &HC0FFFF
            End If
            
            
            '*** SÓ EDITA SE AINDA NÃO FOI CHECADO   ***
            If (UCase(frmLogin.txtUserName) = UCase(NomeMestre) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre2) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre3)) And ADOREG.Recordset.Fields("M_DEM_OK") = 0 Then
                CK_ACF.Enabled = True
                'TXT_AC_F.Enabled = True
            Else
                CK_ACF.Enabled = False
                'TXT_AC_F.Enabled = False
            End If
            
            
         End If
            
            TXT_SIT_EMP = ""
            On Error Resume Next
            TXT_SIT_EMP = de.cnc.Execute("SELECT F_EMPRESTIMO_ANOT FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)

        If adReason <> 7 Then
            'Saldo DO EMPRESTIMO
            If de.rsTAB_FUNCIONARIO.State = 1 Then de.rsTAB_FUNCIONARIO.Requery
            w_Emprest = de.cnc.Execute("Select F_EMPRESTIMO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
            txt_SaldoEmp = IIf(IsNull(w_Emprest), Format(0, "R$ 0.00"), Format(w_Emprest, "R$ 0.00"))
            
            'Dados Contas
            Lancamentos
        End If

    Select Case ADOREG.Recordset.Fields("M_TIPO")
        Case "V": TXT_FTIPO.Caption = "VENDEDOR"
                  TXT_FTIPO.FontSize = 14
        Case "C": TXT_FTIPO.Caption = "CAIXA"
                  TXT_FTIPO.FontSize = 14
        Case "2": TXT_FTIPO.Caption = "2º CAIXA"
                  TXT_FTIPO.FontSize = 12
        Case "X": TXT_FTIPO.Caption = "CX EXTRA"
                  TXT_FTIPO.FontSize = 14
        Case "R": TXT_FTIPO.Caption = "SEGURANÇA"
                  TXT_FTIPO.FontSize = 12
        Case "G": TXT_FTIPO.Caption = "GERENTE"
                  TXT_FTIPO.FontSize = 14
        Case "D": TXT_FTIPO.Caption = "GER RODA"
                  TXT_FTIPO.FontSize = 14
        Case "S": TXT_FTIPO.Caption = "SUPERVISOR"
                  TXT_FTIPO.FontSize = 12
        Case "O": TXT_FTIPO.Caption = "RP"
                  TXT_FTIPO.FontSize = 14
    End Select
    
   Select Case TXT_MES
       Case "1": lblMes.Caption = "Janeiro" & " / " & Right(TXT_ANO, 2)
       Case "2": lblMes.Caption = "Fevereiro" & " / " & Right(TXT_ANO, 2)
       Case "3": lblMes.Caption = "Março" & " / " & Right(TXT_ANO, 2)
       Case "4": lblMes.Caption = "Abril" & " / " & Right(TXT_ANO, 2)
       Case "5": lblMes.Caption = "Maio" & " / " & Right(TXT_ANO, 2)
       Case "6": lblMes.Caption = "Junho" & " / " & Right(TXT_ANO, 2)
       Case "7": lblMes.Caption = "Julho" & " / " & Right(TXT_ANO, 2)
       Case "8": lblMes.Caption = "Agosto" & " / " & Right(TXT_ANO, 2)
       Case "9": lblMes.Caption = "Setembro" & " / " & Right(TXT_ANO, 2)
       Case "10": lblMes.Caption = "Outubro" & " / " & Right(TXT_ANO, 2)
       Case "11": lblMes.Caption = "Novembro" & " / " & Right(TXT_ANO, 2)
       Case "12": lblMes.Caption = "Dezembro" & " / " & Right(TXT_ANO, 2)
    End Select
    
    
    'If adoReg.Recordset.Fields("M_DEM_OK") Then
    '    BarraF.Buttons("nova").Enabled = True
    'Else
    '    BarraF.Buttons("nova").Enabled = False
    'End If
        
    Dim w_Nasc, w_Idade, w_TempoServico As Double
    
    w_Nasc = de.cnc.Execute("SELECT F_DT_NASC FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD")).Fields(0).value
    w_Idade = DateDiff("yyyy", w_Nasc, Date)
    If w_Idade > 1 Then
        txtIdade = FormatNumber(w_Idade, 0) & " anos"
    Else
        txtIdade = "???"
    End If
    
    If IsDate(ADOREG.Recordset.Fields("M_DT_DEM")) Then
        w_TempoServico = DateDiff("m", ADOREG.Recordset.Fields("M_DT_ADM"), ADOREG.Recordset.Fields("M_DT_DEM"))
    Else
        w_TempoServico = DateDiff("m", ADOREG.Recordset.Fields("M_DT_ADM"), Date)
    End If
    If w_TempoServico >= 12 Then
        w_TempoServico = w_TempoServico / 12
        If w_TempoServico = 1 Then
            txtTempoServico = FormatNumber(w_TempoServico, 0) & " ano"
        Else
            If w_TempoServico <= 3 Then
                txtTempoServico = FormatNumber(w_TempoServico, 1) & " anos"
            Else
                txtTempoServico = FormatNumber(w_TempoServico, 0) & " anos"
            End If
        End If
    Else
        Select Case w_TempoServico
            Case 0:
                If IsDate(ADOREG.Recordset.Fields("M_DT_DEM")) Then
                    txtTempoServico = "< 1 mês"
                Else
                    txtTempoServico = "NOVA"
                End If
            Case 1:
                txtTempoServico = FormatNumber(w_TempoServico, 0) & " mês"
            Case 2 To 11:
                txtTempoServico = FormatNumber(w_TempoServico, 0) & " meses"
        End Select
    End If

sair:
    
    V_MOVE = True
    Exit Sub

err1:
    If Err.Number = 3021 Then
        'MsgBox "Registro não encontrado!", vbCritical
    ElseIf Not Err.Number = -2147217885 Then
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    End If
    Resume sair
End Sub

Private Sub ADO_GRID_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1

    If V_MOVE = True Then
        V_MOVE = False
        ADOREG.Recordset.UpdateBatch ' adAffectCurrent
        ADOREG.Refresh
        'adoReg.Recordset.Requery
        ADOREG.Recordset.Move ADO_GRID.Recordset.AbsolutePosition - 1
        V_MOVE = True

    If (ADOREG.Recordset.Fields("M_TIPO") = "V" Or ADOREG.Recordset.Fields("M_TIPO") = "C") Or acessoTotal() Then
            Set ADO_LANC.Recordset = de.cnc.Execute("SELECT TAB_DESC_CALC.C_DATA_INTERNA AS DT_LCTO, 'CT: ' + TAB_TP_CONTA.TP_DESC + '     DESC: ' + TAB_DESC_CALC.C_DESC AS CONTA, TAB_DESC_CALC.C_VALOR AS VALOR, TAB_DESC_CALC.C_TP_OP AS OP FROM TAB_DESC_CALC, TAB_TP_CONTA WHERE TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_DESC_CALC.C_N_FICHA = " & ADO_GRID.Recordset.Fields("M_NFICHA") & ")").Clone
        Else
            Set ADO_LANC.Recordset = de.cnc.Execute("SELECT TAB_DESC_CALC.C_DATA_INTERNA AS DT_LCTO, 'CT: ' + TAB_TP_CONTA.TP_DESC + '     DESC: ' + TAB_DESC_CALC.C_DESC AS CONTA, TAB_DESC_CALC.C_VALOR AS VALOR, TAB_DESC_CALC.C_TP_OP AS OP FROM TAB_DESC_CALC, TAB_TP_CONTA WHERE TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND TAB_TP_CONTA.TP_COD <> 20  AND TAB_TP_CONTA.TP_COD <> 78 AND (TAB_DESC_CALC.C_N_FICHA = " & ADO_GRID.Recordset.Fields("M_NFICHA") & ")").Clone
        End If
        
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
On Error GoTo err1

    Select Case Button.key
        
        Case "emp":
                    If de.rscmdSqlConEmprestimo.State = 1 Then
                        On Error Resume Next
                        de.rscmdSqlConEmprestimo.UpdateBatch adAffectCurrent
                        On Error GoTo err1
                        de.rscmdSqlConEmprestimo.Close
                    End If
                    de.cmdSqlConEmprestimo ADOREG.Recordset.Fields("m_f_cod")
                    
                    If de.rscmdSqlConEmprestimo.RecordCount > 0 Then
                        frm_Emprestimos_Cons.Show 1
                    Else
                        MsgBox "Não existe empréstimo p/ esta Ficha!", vbInformation
                    End If
        Case "fechar": Fechar
        Case "nova":
                        frm_Cad_Funcionario.Show 1
                        'w_Func_atual = ADOREG.Recordset.Fields("M_F_COD")
                        'If (IsNull(de.cnc.Execute("SELECT F_DT_DEM FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & w_Func_atual).Fields(0))) Then
                        '    MsgBox "Ainda existem fichas abertas para o funcionário " & UCase(ADOREG.Recordset.Fields("M_NOME")) & "!", vbCritical, "Não foi possível criar Nova Ficha"
                        'Else
                        '    frm_Cad_Fic_Mensal.Show
                        'End If
        Case "editar": Editar , 1
        Case "salvar": Salvar
        Case "cancelar": Cancelar
        Case "imprimir": Imprimir
        Case "excluir": Excluir
        Case "filtrar": FILTRAR 0
        Case "conta": CONTA
        Case "dupla": VisDupla
        Case "vistar":
                        frm_Alt_Visto_Vale.ckTodas.value = 0
                        frm_Alt_Visto_Vale.TXT_ANO = Me.TXT_ANO
                        frm_Alt_Visto_Vale.TXT_MES = Me.TXT_MES
                        frm_Alt_Visto_Vale.TXT_LOGO = Me.txtLogo
                        frm_Alt_Visto_Vale.ck_Nome.value = 0
                        frm_Alt_Visto_Vale.dbNome.BoundText = Me.TXT_FUNC.BoundText
                        frm_Alt_Visto_Vale.txt_tipo = TXT_FTIPO
                        frm_Alt_Visto_Vale.Show 1
        Case "gcomissao":
                        frm_Gerar_Comissao.ck_Nome.value = 0
                        frm_Gerar_Comissao.TXT_LOGO = ADOREG.Recordset.Fields("M_LOGO")
                        frm_Gerar_Comissao.TXT_MES = ADOREG.Recordset.Fields("M_MES")
                        frm_Gerar_Comissao.TXT_ANO = ADOREG.Recordset.Fields("M_ANO")
                        frm_Gerar_Comissao.dbNome.BoundText = ADOREG.Recordset.Fields("M_F_COD")
                       
                        frm_Gerar_Comissao.Show 1
        Case "programados":
                        w_CodFunc = txt_F_COD
                        frm_Alt_Desc_Calc_fixo.lbFunc.Caption = TXT_FUNC.text
                        frm_Alt_Desc_Calc_fixo.Show 1
        Case "cadastro":
                        frm_Alt_FuncionarioDireto.Show 1
        Case "desbloquear": Desbloquear
        Case "gerente": frm_Opt_GerenteCaixa.Show 1
                                
    End Select
    
    'If Button.key = "nova" Then
    '    W_POS = adoReg.Recordset.AbsolutePosition - 1
    '
    '    If de.rscmdSqlVisAltFichas.State = 1 Then de.rscmdSqlVisAltFichas.Close
    '    de.cmdSqlVisAltFichas txt_PMes, txt_PAno
    '    Set adoReg.Recordset = de.rscmdSqlVisAltFichas.Clone
    '
    '    adoReg.Recordset.Move W_POS
    '
    '    Lancamentos
    '
    'End If
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub



'*** Rotinas ***
Sub Desbloquear()

Dim w_Lojas As String
Dim w_Tipos As String
Dim w_FirstLoja As Boolean
Dim w_FirstTipo As Boolean
Dim w_reg
 
 
On Error GoTo err1
     
    
    If FRM_LIBERAR.ckTodas.value = 0 Then
        FRM_LIBERAR.TXT_LOGO = txtLogo
    Else
        FRM_LIBERAR.TXT_LOGO = "%"
    End If
    
    FRM_LIBERAR.TXT_MES = TXT_MES
    FRM_LIBERAR.TXT_ANO = TXT_ANO
    
    If FRM_LIBERAR.ck_Nome.value = 0 Then
        FRM_LIBERAR.dbNome = TXT_FUNC
    Else
        FRM_LIBERAR.dbNome = "%"
    End If
    FRM_LIBERAR.Show 1
    
    If (FRM_LIBERAR.txt_State = "F") Then
       MsgBox "Ação cancelada!", vbCritical
    Else
        
        
    w_FirstLoja = True
    For I = 0 To FRM_LIBERAR.TXT_LOGO.ListCount - 1
        If FRM_LIBERAR.TXT_LOGO.Selected(I) = True Then
            If w_FirstLoja Then
                w_Lojas = "'" & FRM_LIBERAR.TXT_LOGO.list(I) & "'"
            Else
                w_Lojas = w_Lojas & ",'" & FRM_LIBERAR.TXT_LOGO.list(I) & "'"
            End If
            w_FirstLoja = False
        End If
    Next
    
    'tipos
    w_FirstTipo = True
    Dim w_tipo
    For J = 0 To FRM_LIBERAR.txt_tipo.ListCount - 1
        If FRM_LIBERAR.txt_tipo.Selected(J) = True Then
            Select Case FRM_LIBERAR.txt_tipo.list(J)
                Case "VENDEDOR": w_tipo = "V"
                Case "GERENTE": w_tipo = "G"
                Case "GER RODA": w_tipo = "D"
                Case "CX EXTRA": w_tipo = "X"
                Case "SEGURANÇA": w_tipo = "R"
                Case "CAIXA": w_tipo = "C"
                Case "2º CAIXA": w_tipo = "2"
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
    
        'Se quer varios emp. de uma vez então Filtra p/ não mostra os q/ tem acerto final
        w_ACF = ""
        If (FRM_LIBERAR.dbNome = "%" Or FRM_LIBERAR.ck_Nome = 1) Then
            w_ACF = "and (M_DT_ACF is null or M_DT_ACF ='')"
        End If
        
        w_reg = 0

        Dim w_ado As ADODB.Recordset
        Set w_ado = de.cnc.Execute("SELECT TAB_FUNCIONARIO.F_NOME, TAB_FICHA_MENS.M_NFICHA FROM TAB_FUNCIONARIO INNER JOIN TAB_FICHA_MENS ON TAB_FUNCIONARIO.F_Codigo = TAB_FICHA_MENS.M_F_COD Where (( TAB_FICHA_MENS.M_LOGO IN (" & w_Lojas & ")) and ( TAB_FICHA_MENS.M_TIPO IN (" & w_Tipos & ")) And ((TAB_FICHA_MENS.M_MES) = " & FRM_LIBERAR.TXT_MES & ") And ((TAB_FICHA_MENS.M_ANO) = " & FRM_LIBERAR.TXT_ANO & ") " & IIf(FRM_LIBERAR.dbNome = "%", "", "And TAB_FICHA_MENS.M_NOME Like '" & FRM_LIBERAR.dbNome & "'") & IIf(w_ACF <> "", w_ACF, "") & ") ").Clone
        If FRM_LIBERAR.cbAcao.ListIndex = 0 Then
            de.cnc.Execute "UPDATE  TAB_FUNCIONARIO INNER JOIN TAB_FICHA_MENS ON TAB_FUNCIONARIO.F_Codigo = TAB_FICHA_MENS.M_F_COD SET TAB_FICHA_MENS.M_BLOQ = " & 0 & " Where (( TAB_FICHA_MENS.M_LOGO IN (" & w_Lojas & ")) and ( TAB_FICHA_MENS.M_TIPO IN (" & w_Tipos & ")) And ((TAB_FICHA_MENS.M_MES) = " & FRM_LIBERAR.TXT_MES & ") And ((TAB_FICHA_MENS.M_ANO) = " & FRM_LIBERAR.TXT_ANO & ") " & IIf(FRM_LIBERAR.dbNome = "%", "", "And TAB_FICHA_MENS.M_NOME Like '" & FRM_LIBERAR.dbNome & "'") & IIf(w_ACF <> "", w_ACF, "") & ") ", w_reg
            Do While Not w_ado.EOF
                de.cmdIncluirLog Date, Time, w_usuario, "EDITAR", "FICHA", "FICHA: " & w_ado.Fields(1) & " | FUNCIONÁRIO: " & w_ado.Fields(0) & " | ## LIBERANDO ##"
            w_ado.MoveNext
            Loop
            MsgBox "Foram LIBERADAS " & w_reg & " fichas!", vbInformation
        Else
            de.cnc.Execute "UPDATE  TAB_FUNCIONARIO INNER JOIN TAB_FICHA_MENS ON TAB_FUNCIONARIO.F_Codigo = TAB_FICHA_MENS.M_F_COD SET TAB_FICHA_MENS.M_BLOQ = " & -1 & " Where (( TAB_FICHA_MENS.M_LOGO IN (" & w_Lojas & ")) and ( TAB_FICHA_MENS.M_TIPO IN (" & w_Tipos & ")) And ((TAB_FICHA_MENS.M_MES) = " & FRM_LIBERAR.TXT_MES & ") And ((TAB_FICHA_MENS.M_ANO) = " & FRM_LIBERAR.TXT_ANO & ") " & IIf(FRM_LIBERAR.dbNome = "%", "", "And TAB_FICHA_MENS.M_NOME Like '" & FRM_LIBERAR.dbNome & "'") & IIf(w_ACF <> "", w_ACF, "") & ") ", w_reg
            Do While Not w_ado.EOF
                de.cmdIncluirLog Date, Time, w_usuario, "EDITAR", "FICHA", "FICHA: " & w_ado.Fields(1) & " | FUNCIONÁRIO: " & w_ado.Fields(0) & " | ## BLOQUEANDO ##"
            w_ado.MoveNext
            Loop
            MsgBox "Foram BLOQUEADAS " & w_reg & " fichas!", vbInformation
        End If
        
        
        w_reg = 0
        
        'Cancelar
        'Editar 0


    End If
    
sair:
   ' Salvar
    Exit Sub
err1:
    'If Err.Number <> 3705 And Err.Number <> -2147217864 Then
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    'Set ADO_LANC.Recordset = de.cnc.Execute("SELECT TAB_DESC_CALC.C_DT AS DATA, 'CT: ' + TAB_TP_CONTA.TP_DESC + '     DESC: ' + TAB_DESC_CALC.C_DESC AS CONTA, TAB_DESC_CALC.C_VALOR AS VALOR, TAB_DESC_CALC.C_TP_OP AS OP FROM TAB_DESC_CALC, TAB_TP_CONTA WHERE TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_DESC_CALC.C_N_FICHA = " & frm_Alt_Fic_Mensal_Visualizar.ADOREG.Recordset.Fields("M_NFICHA") & ")").Clone
    'Cancelar
    'Cancelar

    Resume sair
End Sub

Sub VisDupla()
        
        
    frm_Alt_Fic_Mensal_Visualizar_Dupla.Show 1
    
End Sub

Private Sub CONTA()
    w_leitura = 0
 
    If Not isMesValido(txt_F_COD, TXT_MES, TXT_ANO) Then 'Verifica se é mês atual ou passado
        If MsgBox("Você está acessando uma ficha que NÃO É DO MÊS ATUAL. Deseja continuar mesmo assim?", vbYesNo, "Confirmação") = vbNo Then
            Exit Sub
        End If
    End If
     
    If ADOREG.Recordset.Fields("M_BLOQ") Then
        MsgBox "Esta ficha não é do mês atual e está BLOQUEADA! Não será permitido alterações!", vbInformation
        w_leitura = 1
    End If

    frm_Alt_Desc_Calc.lb_form = "mensal"
    frm_Alt_Desc_Calc.LB_FUNC.Caption = TXT_FUNC.text

    frm_Alt_Desc_Calc.Show 1
End Sub

Private Sub Cancelar()
On Error Resume Next

    W_FILTRO = ADOREG.Recordset.Filter
    pos = ADOREG.Recordset.Fields("m_nficha")


    ADOREG.Recordset.CancelBatch adAffectCurrent
    de.rscmdSqlVisAltFichas.Resync
On Error GoTo err1
    If de.rscmdSqlVisAltFichas.State = 1 Then de.rscmdSqlVisAltFichas.Close
    de.cmdSqlVisAltFichas txt_PMes, txt_PAno
    Set ADOREG.Recordset = de.rscmdSqlVisAltFichas.Clone
    
    
    Lancamentos
'    Set ADO_LANC.Recordset = ADOREG.Recordset.Fields("cmdSqlVisAltContas").UnderlyingValue
    
    
    If W_FILTRO <> "0" Then ADOREG.Recordset.Filter = W_FILTRO
    ADOREG.Recordset.MoveFirst
    If W_FILTRO <> "BLOQ <> '###'" Then ADOREG.Recordset.Find "m_nficha = " & pos & ""

    Pause 0.3
    
   
   
   
   Editar 0
    
       '*** DESABILITA O EDITAR ****
   If ADOREG.Recordset.Fields("M_BLOQ") = True Then
        BarraF.Buttons("editar").Enabled = False
   Else
        BarraF.Buttons("editar").Enabled = True
   End If


   
    
sair:
    Exit Sub
err1:
    If Err.Number = 3021 Then
        Form_Load
    Else
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    End If
    Resume sair
End Sub


Private Sub Editar(Optional mesvalido, Optional edicao)
On Error GoTo err1

If IsMissing(mesvalido) Then mesvalido = 1 'Se não passar parametro na chamada, verifica o mes

If IsMissing(edicao) Then edicao = 0

If edicao = 1 Then
    w_ck_vt = ck_pg_vt
    w_ac_f = TXT_AC_F
End If

If mesvalido Then
    If Not isMesValido(txt_F_COD, TXT_MES, TXT_ANO) Then 'Verifica se é mês atual ou passado
        If MsgBox("Você está alterando uma ficha que NÃO É DO MÊS ATUAL. Deseja continuar mesmo assim?", vbYesNo, "Alteração de ficha") = vbNo Then
            Exit Sub
        End If
    If ADOREG.Recordset.Fields("M_BLOQ") Then
        MsgBox "Esta ficha não é do mês atual e está BLOQUEADA!", vbCritical
        Exit Sub
    End If
    End If
End If

    'Desbloqueado ou mestre
    'If adoReg.Recordset.Fields("M_BLOQ") = False Or (UCase(frmLogin.txtUserName) = UCase(NomeMestre) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre2) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre3)) Then
    If ADOREG.Recordset.Fields("M_BLOQ") = False Then
    
        BarraF.Buttons("salvar").Enabled = Not BarraF.Buttons("salvar").Enabled
        BarraF.Buttons("cancelar").Enabled = Not BarraF.Buttons("cancelar").Enabled
        BarraF.Buttons("editar").Enabled = Not BarraF.Buttons("editar").Enabled
        'BarraF.Buttons("nova").Enabled = Not BarraF.Buttons("nova").Enabled
        
        Grid.Enabled = Not Grid.Enabled
        
        txtLogo.Enabled = Not txtLogo.Enabled
        txtLogo2.Enabled = Not txtLogo2.Enabled
        If (UCase(frmLogin.txtUserName) = UCase(NomeMestre) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre2) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre3)) Then txt_DT_ADM.Enabled = Not txt_DT_ADM.Enabled
        TXT_DT_REG.Enabled = Not TXT_DT_REG.Enabled
        
        '*** SÓ EDITA SE AINDA NÃO FOI CHECADO   ***
        If (CK_DEM.value = 1 And (UCase(frmLogin.txtUserName) = UCase(NomeMestre) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre2) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre3))) Or (CK_DEM.value = 0) Then
            TXT_DT_DEM.Enabled = Not TXT_DT_DEM.Enabled
        End If
        
        'Permite alterar se for MASTER  ou se ainda não foi checado
        If (CK_13.value = 1 And (UCase(frmLogin.txtUserName) = UCase(NomeMestre) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre2) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre3))) Or CK_13.value = 0 Then
            TXT_13_PG.Enabled = Not TXT_13_PG.Enabled
            TXT_13_ULT_PG.Enabled = Not TXT_13_ULT_PG.Enabled
            TXT_13_OBS.Enabled = Not TXT_13_OBS.Enabled
        End If
        
        'Permite alterar se for MASTER  ou se ainda não foi checado
        If (UCase(frmLogin.txtUserName) = UCase(NomeMestre) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre2) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre3)) Or CK_FERIAS.value = 0 Then
            TXT_FERIAS_PG.Enabled = Not TXT_FERIAS_PG.Enabled
            TXT_FERIAS_ULT_PG.Enabled = Not TXT_FERIAS_ULT_PG.Enabled
        End If
        
        If (UCase(frmLogin.txtUserName) = UCase(NomeMestre) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre2) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre3)) Then
            CK_FERIAS.Enabled = Not CK_FERIAS.Enabled
            CK_13.Enabled = Not CK_13.Enabled
        End If
        
        ck_Acordo.Enabled = Not ck_Acordo.Enabled
        
        TXT_AC_F.Locked = Not TXT_AC_F.Locked
        TXT_FERIAS.Locked = Not TXT_FERIAS.Locked
        
        TXT_OBS.Locked = Not TXT_OBS.Locked
        txt_notas.Locked = Not txt_notas.Locked
        TXT_ANOTACAO.Locked = Not TXT_ANOTACAO.Locked
        txt_Vcto_ferias.Enabled = Not txt_Vcto_ferias.Enabled
        
        txt_NFilhos.Enabled = Not txt_NFilhos.Enabled
        ck_pg_SFam.Enabled = Not ck_pg_SFam.Enabled
        
        ck_pg_vt.Enabled = Not ck_pg_vt.Enabled
        
    On Error Resume Next
        If BarraF.Buttons("salvar").Enabled = False And Grid.Enabled = True Then
            Grid.SetFocus
        'ElseIf TXT_DT_REG.Enabled = True Then
            'TXT_DT_REG.SetFocus
    '        txt_DT_ADM.SetFocus
        End If
    
        If BarraF.Buttons("salvar").Enabled = True Then
            Frame1.Enabled = False
            w_logo = txtLogo
        Else
            Frame1.Enabled = True
        End If
    
    Else
        If edicao = 0 Then
            On Error Resume Next
        Else
            If TXT_AC_F = "" Then
                MsgBox "Não é possível alterar uma ficha anterior ao mês passado!", vbExclamation
            Else
                MsgBox "Não é possível alterar uma ficha que já foi feito acerto final!", vbExclamation
            End If
        End If
    End If
sair:
    Exit Sub
err1:

    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Excluir()
On Error GoTo err1
        
        
    If vbYes = MsgBox("DESEJA REALMENTE EXCLUIR A FICHA MENSAL (" & TXT_NFICHA & " : " & TXT_FUNC & ")?" & vbNewLine & vbNewLine & "VOCÊ ESTÁ EXCLUINDO A FICHA E NÃO O LANÇAMENTO.", vbQuestion + vbYesNo) Then
    frm_Habilitar.Show 1
    w_PSS = frm_Habilitar.txt_Pss



If w_PSS = w_PassWordLib Then
        
          
        
If Not isMesValido(txt_F_COD, TXT_MES, TXT_ANO) Then 'Verifica se é mês atual ou passado
    If MsgBox("Você está excluindo uma ficha que NÃO É DO MÊS ATUAL. Deseja continuar mesmo assim?", vbYesNo, "Exclusão de ficha") = vbNo Then
        Exit Sub
    End If
    If ADOREG.Recordset.Fields("M_BLOQ") Then
        MsgBox "Esta ficha não é do mês atual e está BLOQUEADA!", vbCritical
        Exit Sub
    End If
End If



        W_POS = ADOREG.Recordset.AbsolutePosition - 1
        de.cnc.Execute "DELETE * FROM TAB_FICHA_MENS WHERE M_NFICHA = " & TXT_NFICHA & "", REG_AF
        If REG_AF = 1 Then
            MsgBox "Registro excluído com sucesso!", vbInformation
        Else
            MsgBox "Não foi possível excluir!", vbCritical
        End If
    '        ADOREG.Recordset.Delete adAffectCurrent
        
    
    On Error Resume Next
    'Set adoReg.Recordset = Nothing
    'Set ADO_LANC.Recordset = Nothing
    'Set adoReg.Recordset = Nothing
    'Set ADO_LANC.Recordset = Nothing
    
    
    de.rscmdSqlVisAltFichas.Close
    de.cmdSqlVisAltFichas TXT_MES, TXT_ANO
    Set ADOREG.Recordset = de.rscmdSqlVisAltFichas.Clone
    
    If (ADOREG.Recordset.Fields("M_TIPO") = "V" Or ADOREG.Recordset.Fields("M_TIPO") = "C") Or acessoTotal() Then
        Set ADO_LANC.Recordset = ADOREG.Recordset.Fields("cmdSqlVisAltContas3").value
    Else
        Set ADO_LANC.Recordset = ADOREG.Recordset.Fields("cmdSqlVisAltContas2").value
    End If
    
    Cancelar
    Cancelar
        w_PSS = ""
        

Else
    MsgBox "Senha de Liberação Incorreta!", vbCritical

End If
 
    End If
 
 
 
sair:
    Exit Sub
err1:
    If Not Err.Number = -2147467259 Then
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    Else
        MsgBox "NÃO É POSSÍVEL EXCLUIR ESTA FICHA MENSAL, DEVIDO A CÁLCULOS RELACIONADAS A ELA!", vbCritical
        ADOREG.Refresh
    End If
    Resume sair
End Sub

Private Sub Fechar()
On Error GoTo err1
    'If de.rsTAB_FICHA_MENS.State = 1 Then

    '    de.rsTAB_FICHA_MENS.Close
    '    de.TAB_FICHA_MENS
        
    'End If
    'de.rsTAB_DESC_CALC.Requery
    'Hide
sair:
    Unload Me
    Exit Sub
err1:
    'MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub




Private Sub Salvar()
On Error GoTo err1
    
    If w_SN_Total = False Then GoTo sair
    ADOREG.Recordset.UpdateBatch adAffectCurrent
        
    '*** Atualiza o Funcionário ****
    
    'If Trim(TXT_AC_F) = "" Then
    '    de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_DT_ACF = NULL WHERE (M_NFICHA = " & TXT_NFICHA & " )", w_reg
    '    adoReg.Recordset.Fields("M_DT_ACF") = Null
    'End If
    
    
    If txt_DT_ADM = "" Then
        MsgBox "Data de ADM  não pode estar em branco, será inserida a data do dia automaticamente!", vbCritical
        txt_DT_ADM = Date
    End If
      
        w_dt_REg = IIf(TXT_DT_REG = "", Null, Format(TXT_DT_REG, "DD/MM/YYYY"))
        
        If IsNull(w_dt_REg) Then
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS = '" & TXT_FERIAS & "', F_OBS = '" & TXT_OBS & "', F_NOTAS = '" & txt_notas & "', F_ANOTACAO = '" & TXT_ANOTACAO & "', F_DT_ADM = '" & Format(txt_DT_ADM, "DD/MM/YYYY") & "', F_DT_REG = NULL, F_13_OBS = '" & TXT_13_OBS & "'  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        Else
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS = '" & TXT_FERIAS & "', F_OBS = '" & TXT_OBS & "', F_NOTAS = '" & txt_notas & "', F_ANOTACAO = '" & TXT_ANOTACAO & "', F_DT_ADM = '" & Format(txt_DT_ADM, "DD/MM/YYYY") & "', F_DT_REG = '" & w_dt_REg & "', F_13_OBS = '" & TXT_13_OBS & "'  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        End If
        
        If w_logo <> txtLogo Then
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_COD_L = '" & txtLogo & "'   WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        End If
        
        'ATUALIZA DATA DE DEMISSÃO DO CADASTRO
        w_dt_DEM = IIf(TXT_DT_DEM = "", Null, Format(TXT_DT_DEM, "DD/MM/YYYY"))
        'Se não tiver fichas abertas
        If de.cnc.Execute("SELECT M_NFICHA FROM TAB_FICHA_MENS WHERE M_DT_DEM IS NULL AND M_F_COD = " & txt_F_COD & " AND ((M_MES >= " & TXT_MES & " AND M_ANO = " & TXT_ANO & ") OR M_ANO > " & TXT_ANO & ")").RecordCount = 0 Then
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_DT_DEM = '" & w_dt_DEM & "'  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        Else
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_DT_DEM = NULL  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        End If

        w_reg = 0
        '*** OK F_FERIAS_OK  , 13_OK, DT_DEM_OK
        de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS_OK = " & CInt(ADOREG.Recordset.Fields("M_FERIAS_OK")) & _
                                                ", F_13_OK = " & CInt(ADOREG.Recordset.Fields("M_13_OK")) & _
                                                ", F_NUM_FILHOS = " & ADOREG.Recordset.Fields("M_NUM_FILHOS") & _
                                                ", F_DEM_OK = " & CInt(ADOREG.Recordset.Fields("M_DEM_OK")) & _
                                                ", F_PG_SAL_FAM = " & CInt(ADOREG.Recordset.Fields("M_PG_SAL_FAM")) & _
        "  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        
        'e 13_OK ***
        'de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_13_OK = " & ADOREG.Recordset.Fields("M_13_OK") & "  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        '*** OK DT_DEM
        'de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_DEM_OK = " & ADOREG.Recordset.Fields("M_DEM_OK") & "  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
                
                
        If TXT_FERIAS_PG = "" Then
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS_PG = NULL  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        Else
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS_PG = '" & Format(TXT_FERIAS_PG, "DD/MM/YYYY") & "'  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        End If
        
        If TXT_FERIAS_ULT_PG = "" Then
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS_ULT_PG = NULL  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        Else
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS_ULT_PG = '" & Format(TXT_FERIAS_ULT_PG, "DD/MM/YYYY") & "'  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        End If
    
        If TXT_13_PG = "" Then
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_13_PG = NULL WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        Else
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_13_PG = '" & Format(TXT_13_PG, "DD/MM/YYYY") & "'  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        End If
        If TXT_13_ULT_PG = "" Then
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_13_ULT_PG = NULL WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        Else
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_13_ULT_PG = '" & Format(TXT_13_ULT_PG, "DD/MM/YYYY") & "'  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        End If
        If txt_Vcto_ferias <> "" Then
            'ATUALIZA     VCTO DE FERIAS
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_VCTO_FERIAS = '" & txt_Vcto_ferias & "'  WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
        End If
        
        'Paga Vale Transporte? Atualizada cadastro e gera programados
        If (ck_pg_vt = 1 And w_ck_vt = False) Or (ck_pg_vt = 0 And w_ck_vt = True) Then
            
            

            
            If ck_pg_vt Then
            
                If (de.cnc.Execute("SELECT * FROM TAB_DESC_CALC WHERE C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & " AND (C_TP_CONTA = 109 OR C_TP_CONTA = 110 OR C_TP_CONTA = 111)").EOF) Then
            
                    de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_PG_VT = 1 WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
                    de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & " AND (C_TP_CONTA = 109 OR C_TP_CONTA = 110 OR C_TP_CONTA = 111)")
                    de.cnc.Execute ("DELETE FROM TAB_DESC_CALC_FIXO WHERE CF_EMP_COD = " & txt_F_COD & " AND (CF_TP_CONTA = 109 OR CF_TP_CONTA = 110 OR CF_TP_CONTA = 111)")
                        
                    Dim adoFixos As ADODB.Recordset
                    
                    Dim fichaAtual As String
                    Dim ultimoFixo As String
                
                    fichaAtual = ADOREG.Recordset.Fields("M_NFICHA")
                    
                    de.cmdIncluirDescCalcFixo Now(), txt_F_COD, "109", "-", "0", "INSS 8% do piso [GERADO AUTOMATICAMENTE]"
                    ultimoFixo = de.cnc.Execute("SELECT Max([CF_CODIGO]) FROM TAB_DESC_CALC_FIXO").Fields(0)
                    Set adoFixos = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_CODIGO = " & ultimoFixo).Clone
                    de.cmdIncluirDescCalc2 Date, fichaAtual, adoFixos.Fields("CF_TP_CONTA"), adoFixos.Fields("CF_TP_OP"), adoFixos.Fields("CF_VALOR"), adoFixos.Fields("CF_DESC"), "0", adoFixos.Fields("CF_CODIGO"), "0", "0", adoFixos.Fields("CF_EMP_COD"), 0
    
                    ultimoFixo = Empty
                    Set adoFixos = Nothing
                
                    de.cmdIncluirDescCalcFixo Now(), txt_F_COD, "110", "-", "0", "Vale Transporte 6% do piso [GERADO AUTOMATICAMENTE]"
                    ultimoFixo = de.cnc.Execute("SELECT Max([CF_CODIGO]) FROM TAB_DESC_CALC_FIXO").Fields(0)
                    Set adoFixos = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_CODIGO = " & ultimoFixo).Clone
                    de.cmdIncluirDescCalc2 Date, fichaAtual, adoFixos.Fields("CF_TP_CONTA"), adoFixos.Fields("CF_TP_OP"), adoFixos.Fields("CF_VALOR"), adoFixos.Fields("CF_DESC"), "0", adoFixos.Fields("CF_CODIGO"), "0", "0", adoFixos.Fields("CF_EMP_COD"), 0
                
                    ultimoFixo = Empty
                    Set adoFixos = Nothing
                
                    de.cmdIncluirDescCalcFixo Now(), txt_F_COD, "111", "=", "0", "Pagto. de passes (vale transporte) [GERADO AUTOMATICAMENTE]"
                    ultimoFixo = de.cnc.Execute("SELECT Max([CF_CODIGO]) FROM TAB_DESC_CALC_FIXO").Fields(0)
                    Set adoFixos = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_CODIGO = " & ultimoFixo).Clone
                    de.cmdIncluirDescCalc2 Date, fichaAtual, adoFixos.Fields("CF_TP_CONTA"), adoFixos.Fields("CF_TP_OP"), adoFixos.Fields("CF_VALOR"), adoFixos.Fields("CF_DESC"), "0", adoFixos.Fields("CF_CODIGO"), "0", "0", adoFixos.Fields("CF_EMP_COD"), 0
                
                    fichaAtual = Empty
                    ultimoFixo = Empty
                    Set adoFixos = Nothing
                    
                    Lancamentos
                 End If
            Else
                'If vbNo = MsgBox("Você está removendo o Vale Transporte do mês e dos programados." & vbNewLine & "Tem certeza que deseja removê-lo?", vbQuestion + vbYesNo, "Confirmação de Remoção de VT") Then
                '    ck_pg_vt = 1
                '    Exit Sub
                'Else
                    de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_PG_VT = 0 WHERE (F_Codigo = " & txt_F_COD & " )", w_reg
                    de.cnc.Execute ("DELETE FROM TAB_DESC_CALC WHERE C_N_FICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & " AND (C_TP_CONTA = 109 OR C_TP_CONTA = 110 OR C_TP_CONTA = 111)")
                    de.cnc.Execute ("DELETE FROM TAB_DESC_CALC_FIXO WHERE CF_EMP_COD  = " & txt_F_COD & " AND (CF_TP_CONTA = 109 OR CF_TP_CONTA = 110 OR CF_TP_CONTA = 111)")
                    
                    de.cmdIncluirLog Date, Time, w_usuario, "EXCLUIR", "VALE TRANSPORTE", "FICHA: " & TXT_NFICHA & " | FUNCIONÁRIO: " & txt_F_COD & " - " & TXT_FUNC
                    
                    Lancamentos

                'End If
        
            End If
        End If
        
        
    If w_reg = 0 Then MsgBox "Não foi possível atualizar o cadastro de funcionários (as férias e observações)", vbCritical
    
    Editar 0
    
    'Carimbo
    If ADOREG.Recordset.Fields("M_DT_ACF") <> "" Then
        CARIMBO.Visible = True
    Else
        CARIMBO.Visible = False
    End If

    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
    
End Sub


Private Sub Grid_Error(ByVal DataError As Integer, Response As Integer)

   Response = 0
   Timer1.Enabled = True
End Sub

Private Sub GRID_L_DblClick()
    CONTA
End Sub


Private Sub GRID_L_Error(ByVal DataError As Integer, Response As Integer)
    MsgBox DataError & " : " & Response
End Sub

Private Sub GRID_L_GotFocus()
    mnu.Enabled = True
End Sub

Private Sub GRID_L_LostFocus()
    mnu.Enabled = False
End Sub


Private Sub GRID_L_KeyDown(KeyCode As Integer, Shift As Integer)
        
        
        If (UCase(frmLogin.txtUserName) = UCase(NomeMestre) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre2) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre3)) And Shift = 0 And KeyCode <> 13 Then
            'F7
           Select Case KeyCode
            Case 115: mnuAcessoTotal_Click 'F4
            Case 118: mnuVis_Click  'F7
            Case 119: mnuRem_Click  'F8
            Case 122: mnuVist_Click 'F11
            Case 123: mnuRemT_Click 'F12
          End Select
        ElseIf Shift <> 2 And KeyCode = 13 Then
            If Grid.Enabled = True Then
                Grid.SetFocus
            Else
                txt_DT_ADM.SetFocus
            End If
        End If
End Sub




Private Sub GRID_L_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 And (UCase(frmLogin.txtUserName) = UCase(NomeMestre) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre2) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre3)) And CK_ACF = 0 Then
        PopupMenu mnu
    End If
End Sub

Private Sub MSHFlexGrid1_Click()

End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)

     Select Case KeyCode
      Case 115: mnuAcessoTotal_Click 'F4
     End Select
    
End Sub

Sub Op_Click(Index As Integer)
On Error Resume Next

'cbMostrar.SelText = "TODOS"

  W_INDEX = Index
  cmdFiltrar.Visible = True
  TXT_AC_F_Modelo.Visible = False
  
  If Index = 2 Then
     p_Dg.Visible = False
     p_MA.Visible = True
     txt_PMes.SetFocus
  ElseIf Index = 5 Then
     p_Dg.Visible = False
     p_MA.Visible = False
     
     W_LD_FILTRO = False
     ADOREG.Recordset.Filter = 0
     cbMostrar.ListIndex = 0

     Grid.Height = 6915
     frmQtde.Visible = False
      
  ElseIf Index = 6 Then
    
     p_Dg.Visible = False
     p_MA.Visible = False
     
     FILTRAR W_INDEX

    If (UCase(w_usuario) = UCase(NomeMestre) Or UCase(w_usuario) = UCase(NomeMestre2) Or UCase(w_usuario) = UCase(NomeMestre3)) Then
        cmdFiltrar.Visible = False
        TXT_AC_F_Modelo.Visible = True
        TXT_AC_F_Modelo = Format(Date, "DD/MM/YYYY") & "  : "
    End If
  ElseIf Index = 7 Then  'VCTO FERIAS
    
     p_Dg.Visible = False
     p_MA.Visible = False
     
     FILTRAR W_INDEX
  
  
  ElseIf Index = 4 Then
     p_Dg.Visible = False
     p_MA.Visible = False
     FILTRAR W_INDEX
            
            
  Else
     p_Dg.Visible = True
     p_MA.Visible = False
     txt_Pesq.SetFocus
  End If


     If Index <> 4 And Index <> 6 And Index <> 7 Then
        W_CK_DEM = False
        Set ADOREG.Recordset = de.rscmdSqlVisAltFichas.Clone
        
        W_CK_DEM = True
     End If
      'Dados Contas
    Lancamentos

End Sub



Private Sub optLoja_Click()
    Dim w_cod_atual As String
    
    w_cod_atual = txt_F_COD
        
    ADOREG.Recordset.Sort = "B"
    cmdFiltrar_Click
        
    'ADOREG.Recordset.MoveFirst
    'ADOREG.Recordset.Find "m_f_cod = " & w_cod_atual, , adSearchForward
End Sub

Private Sub optNome_Click()
    Dim w_cod_atual As String
    
    w_cod_atual = txt_F_COD
    
    ADOREG.Recordset.Sort = "M_NOME"
    cmdFiltrar_Click
        
    'ADOREG.Recordset.MoveFirst
    'ADOREG.Recordset.Find "m_f_cod = " & w_cod_atual, , adSearchForward
    
End Sub

Private Sub optTipo_Click()
    Dim w_cod_atual As String
    
    w_cod_atual = txt_F_COD
    
    ADOREG.Recordset.Sort = "M_TIPO"
    cmdFiltrar_Click
        
    'ADOREG.Recordset.MoveFirst
    'ADOREG.Recordset.Find "m_f_cod = " & w_cod_atual, , adSearchForward
End Sub

Private Sub Text1_GotFocus()
    With Text1
         'Seleciona tudo
         .SelStart = 0
         .SelLength = Len(Text1.text)
         .SetFocus
         
        ' Posiciona o cursor no fim do texto
        '.SelStart = Len(.Text)
    End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    ' ao pressionar ENTER aceitar a entrada de dados
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        AtribuiValorCelula
        'ProximaCelula
    ' ESC, cancela a edição
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Text1.Visible = False
        ControlVisible = False
    End If
End Sub

Private Sub Text1_LostFocus()
    OcultarControles
End Sub

Private Sub Text2_Change()

End Sub

Private Sub Timer1_Timer()
    Form_Activate
    Timer1.Enabled = False
End Sub

Private Sub TXT_13_OBS_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub TXT_AC_F_Change()
    If Not ADOREG.Recordset.EOF Then
        If TXT_AC_F <> w_ac_f Then
           If w_SN_Total = True And txt_F_COD = ADOREG.Recordset.Fields("M_F_COD") And (UCase(frmLogin.txtUserName) = UCase(NomeMestre) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre2) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre3)) And TXT_NFICHA = ADOREG.Recordset.Fields("M_NFICHA") And BarraF.Buttons("salvar").Enabled = False Then
                Editar 0
                If TXT_AC_F = Null Or Trim(TXT_AC_F) = "" Then
                    'TXT_AC_F = Format(Date, "DD/MM/YYYY") & "  : "
                
    '           If TXT_AC_F.Enabled = True Then
                    TXT_AC_F.SetFocus
                    Sendkeys "{END}"
                End If
            End If
        End If 'referente ao If TXT_AC_F <> w_ac_f Then
     End If
End Sub

Private Sub TXT_AC_F_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
End Sub

Private Sub TXT_AC_F_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
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
'    KeyEnter KeyCode
End Sub

Private Sub txt_DT_ADM_Validate(Cancel As Boolean)
'    If TXT_DT_REG = "" Then txt_Vcto_ferias = Format(txt_DT_ADM, "MM")
End Sub

Private Sub TXT_DT_REG_KeyDown(KeyCode As Integer, Shift As Integer)
'      If Shift <> 2 And KeyCode = 13 Then SendKeys "{tab}"
End Sub
Private Sub C_KeyDown(KeyCode As Integer, Shift As Integer)
'      If Shift <> 2 And KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub TXT_DT_REG_Validate(Cancel As Boolean)
   ' If IsDate(TXT_DT_REG) Then txt_Vcto_ferias = Format(TXT_DT_REG, "MM")
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

Private Sub TXT_LOGO2_Change()
    'txtLogo.BoundText = TXT_LOGO2.BoundText
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
            FILTRAR 0
    Case 79: ' "O"
            CONTA
    End Select
ElseIf KeyCode = 116 And Shift = 0 And w_F5 = False Then
     If BarraF.Buttons("dupla").Visible = True Then
        w_F5 = True
        VisDupla
     End If
ElseIf w_F5 = True Then
    w_F5 = False
End If
End Sub

Private Sub txt_PAno_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then
        cmdFiltrar.SetFocus
        FILTRAR W_INDEX
      End If
End Sub

Private Sub txt_Pesq_Change()
    If Op(0).value = False Then
        
        'cmdFiltrar_Click
        
    End If
End Sub

Private Sub txt_Pesq_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then
        cmdFiltrar.SetFocus
        cmdFiltrar_Click
      End If
      
     Select Case KeyCode
        Case 115: mnuAcessoTotal_Click 'F4
     End Select
End Sub

Private Sub txt_PMes_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then KeyEnter (KeyCode)
End Sub




Private Sub txt_Vcto_ferias_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then Sendkeys "{tab}"

End Sub

Private Sub txtLogo_Click(Area As Integer)
    txtLogo2.BoundText = txtLogo.BoundText
End Sub

Private Sub txtLogo_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub


Private Sub FILTRAR(Index As Byte)
Dim w_resp As String
Dim W_CAMPO As String

On Error GoTo err1
    
    w_resp = Index + 1

    If Not w_resp = "" And IsNumeric(w_resp) And w_resp >= 1 And w_resp <= 5 Or w_resp = 7 Or w_resp = 8 Or w_resp = 9 Or w_resp = 10 Or w_resp = 11 Then
        Select Case w_resp
        'Nº
        Case 1:
            w_resp = "Nº"
            W_CAMPO = "M_F_Cod"
        'LOGO
        Case 2:
            w_resp = "LOGO"
            W_CAMPO = "B"
        'MÊS/ANO
        Case 3:
            w_resp = "MÊS / ANO"
            W_CAMPO = "M_MES"
            W_CAMPO2 = "M_ANO"
        'EMP
        Case 4:
            w_resp = "EMP."
            W_CAMPO = "M_NOME"
        'Saldo de emprestimo
        Case 5:
            Dim w_ado As ADODB.Recordset
            
            Set w_ado = de.cnc.Execute("SELECT F_Codigo, F_EMPRESTIMO FROM TAB_FUNCIONARIO WHERE (F_EMPRESTIMO > 0)").Clone
          
            W_FILTRO = ""
            Do While Not w_ado.EOF
                W_FILTRO = W_FILTRO & IIf(Len(W_FILTRO) > 0, " or ", "") & "M_F_COD = " & w_ado.Fields(0)
                        
                w_ado.MoveNext
            Loop
            
            '*** filtrar ***
            W_LD_FILTRO = True
            
            ADOREG.Recordset.Filter = W_FILTRO
            If W_FILTRO = "" Then MsgBox "NÃO EXISTE FICHA COM EMPRÉSTIMOS!", vbInformation
        Case 7:

            Dim w_Ado1 As ADODB.Recordset
            
            'Set w_Ado1 = de.cnc.Execute("SELECT M_NFICHA, M_DT_DEM FROM TAB_FICHA_MENS WHERE NOT (M_DT_DEM IS NULL) AND M_MES = " & ADOREG.Recordset.Fields("M_MES") & " AND M_ANO = " & ADOREG.Recordset.Fields("M_ANO") & "").Clone
            Set w_Ado1 = de.cnc.Execute("SELECT M_NFICHA, M_DT_DEM FROM TAB_FICHA_MENS WHERE NOT (M_DT_DEM IS NULL) AND M_ANO = " & ADOREG.Recordset.Fields("M_ANO") & "").Clone
          
            W_FILTRO = ""
            Do While Not w_Ado1.EOF
                W_FILTRO = W_FILTRO & IIf(Len(W_FILTRO) > 0, " or ", "") & "M_NFICHA = " & w_Ado1.Fields("M_NFICHA")
                w_Ado1.MoveNext
            Loop
            
            '*** filtrar ***
            W_LD_FILTRO = True
            
            ADOREG.Recordset.Filter = W_FILTRO
            If W_FILTRO = "" Then MsgBox "NÃO EXISTE NENHUM EMP. (D)!", vbInformation
        
        Case 8: 'VCTO DE FERIAS
            
            
            Dim w_Ado2 As ADODB.Recordset
            Dim W_INCLUIR_FILTRO As Boolean
            
            Set w_Ado2 = de.cnc.Execute("SELECT M_NFICHA, M_FERIAS_PG, M_DT_REG, M_DT_ADM FROM TAB_FICHA_MENS WHERE (M_DT_DEM IS NULL) AND M_MES = " & ADOREG.Recordset.Fields("M_MES") & " AND M_ANO = " & ADOREG.Recordset.Fields("M_ANO") & " AND M_VCTO_FERIAS = " & TXT_MES & "").Clone
          
            W_FILTRO = ""
            Do While Not w_Ado2.EOF
                W_INCLUIR_FILTRO = False
                
                'If IsNull(w_Ado2.Fields("M_FERIAS_PG")) Then
                '    If Not IsNull(w_Ado2.Fields("M_DT_REG")) Then
                '        If Year(w_Ado2.Fields("M_DT_REG")) < TXT_ANO Then
                '            W_INCLUIR_FILTRO = True
                '        Else
                '            W_INCLUIR_FILTRO = False
                '        End If
                '    ElseIf Not IsNull(w_Ado2.Fields("M_DT_ADM")) Then
                '        If Year(w_Ado2.Fields("M_DT_ADM")) < TXT_ANO Then
                '            W_INCLUIR_FILTRO = True
                '        Else
                '            W_INCLUIR_FILTRO = False
                '        End If
                 '   End If
                'Else
                '    If Year(w_Ado2.Fields("M_FERIAS_PG")) < TXT_ANO Then
                '        W_INCLUIR_FILTRO = True
                '    Else
                '        W_INCLUIR_FILTRO = False
                '    End If
                'End If
               '
          '****************************
                    If Not IsNull(w_Ado2.Fields("M_DT_REG")) Then
                        If Year(w_Ado2.Fields("M_DT_REG")) < TXT_ANO Then
                            W_INCLUIR_FILTRO = True
                        Else
                            W_INCLUIR_FILTRO = False
                        End If
                    ElseIf Not IsNull(w_Ado2.Fields("M_DT_ADM")) Then
                        If Year(w_Ado2.Fields("M_DT_ADM")) < TXT_ANO Then
                            W_INCLUIR_FILTRO = True
                        Else
                            W_INCLUIR_FILTRO = False
                        End If
                    End If
                
                
                
                If W_INCLUIR_FILTRO = True Then W_FILTRO = W_FILTRO & IIf(Len(W_FILTRO) > 0, " or ", "") & "M_NFICHA = " & w_Ado2.Fields("M_NFICHA")
                
                w_Ado2.MoveNext
            Loop
            
            '*** filtrar ***
            W_LD_FILTRO = True
            
            ADOREG.Recordset.Filter = W_FILTRO
            'If W_FILTRO = "" Then
            '    Op(5).Value = True
            '    MsgBox "NÃO EXISTE NENHUM EMP. COM (F) VENCIDA, QUE AINDA NÃO FOI PAGA!", vbInformation
            'End If
        '*** REMOVE O FILTRO ****
        Case 6:
            If Not ADOREG.Recordset.Filter = 0 Then
                W_LD_FILTRO = False
                ADOREG.Recordset.Filter = 0
                de.rscmdSqlVisAltFichas.Requery
                'Set ADOREG.Recordset = de.rscmdSqlVisAltFichas.Clone
                
                Lancamentos
                
'                ADOREG.Refresh
'                ADO_LANC.Refresh
            End If
        Case 9:
            w_resp = "TIPO_GERENTE"
            W_CAMPO = "M_TIPO"
        Case 10:
            w_resp = "TIPO_OUTROS"
            W_CAMPO = "M_TIPO"
        Case 11:
            w_resp = "NAO_BLOQUEADOS"
            W_CAMPO = "BLOQ"
        End Select
        
        '*** Não filtra qdo for 6  ou 5
        If Not w_resp = "6" And Not w_resp = "5" And Not w_resp = "7" And Not w_resp = "8" Then
            If w_resp = "Nº" Then
                W_FILTRO = W_CAMPO & " = " & txt_Pesq
                W_LD_FILTRO = True
                ADOREG.Recordset.Filter = W_FILTRO
                
            ElseIf w_resp = "TIPO_GERENTE" Then
                W_FILTRO = W_CAMPO & " = " & txt_Pesq
                W_LD_FILTRO = True
                ADOREG.Recordset.Filter = W_FILTRO
                
            ElseIf w_resp = "TIPO_OUTROS" Or w_resp = "NAO_BLOQUEADOS" Then
                W_FILTRO = W_CAMPO & " <> " & txt_Pesq
                W_LD_FILTRO = True
                ADOREG.Recordset.Filter = W_FILTRO
                
            ElseIf w_resp = "LOGO" Or w_resp = "EMP." Then
                W_FILTRO = W_CAMPO & " LIKE '%" & txt_Pesq & "%'"
                W_LD_FILTRO = True
                ADOREG.Recordset.Filter = W_FILTRO
            
            Lancamentos
           
            Else
                W_FILTRO = txt_PMes
                W_FILTRO1 = txt_PAno
                
                If Not W_FILTRO = "" And IsNumeric(W_FILTRO) And IsNumeric(W_FILTRO1) And Len(W_FILTRO1) = 4 Then
                    'On Error Resume Next
                    de.rscmdSqlVisAltFichas.CancelBatch
                    If de.rscmdSqlVisAltFichas.State = 1 Then de.rscmdSqlVisAltFichas.Close
                    de.cmdSqlVisAltFichas W_FILTRO, W_FILTRO1
                    Set ADOREG.Recordset = de.rscmdSqlVisAltFichas.Clone
                    
                Lancamentos
                    
                    W_LD_FILTRO = True
                End If
                                   
            End If
        End If
        If ADOREG.Recordset.RecordCount <= 0 Then
            MsgBox "Não existe ficha com a descrição solicitada!", vbExclamation
                W_LD_FILTRO = False
                ADOREG.Recordset.Filter = 0
                Set ADOREG.Recordset = de.rscmdSqlVisAltFichas.Clone
                
        End If
            
        'Saldo DO EMPRESTIMO
        If de.rsTAB_FUNCIONARIO.State = 1 Then de.rsTAB_FUNCIONARIO.Requery
        w_Emprest = de.cnc.Execute("Select F_EMPRESTIMO FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
        txt_SaldoEmp = IIf(IsNull(w_Emprest), Format(0, "R$ 0.00"), Format(w_Emprest, "R$ 0.00"))
             
    End If
    
    If Index = 1 Then
        Grid.Height = 4995
        frmQtde.Visible = True
        Dim adoQtde As ADODB.Recordset
        
        Set adoQtde = ADOREG.Recordset.Clone
        
        txtQtdeG = 0
        txtQtdeV = 0
        txtQtdeC = 0
        txtQtdeX = 0
        txtQtdeR = 0
        txtQtdeDEM = 0
        
        adoQtde.Filter = W_FILTRO
        
        adoQtde.MoveFirst
        txtQtdeLimiteV = IIf(IsNull(de.cnc.Execute("SELECT QtdeLimiteVend FROM lojb010 WHERE COD_LOJ = '" & adoQtde.Fields("M_LOGO") & "'").Fields(0)), 0, de.cnc.Execute("SELECT QtdeLimiteVend FROM lojb010 WHERE COD_LOJ = '" & adoQtde.Fields("M_LOGO") & "'").Fields(0))
        wTxtOld = txtQtdeLimiteV
        Do While Not adoQtde.EOF
            If IsNull(adoQtde.Fields("M_DT_DEM")) Then
                Select Case adoQtde.Fields("M_TIPO")
                    Case "G": txtQtdeG = CInt(txtQtdeG) + 1
                    Case "V": txtQtdeV = CInt(txtQtdeV) + 1
                    Case "C": txtQtdeC = CInt(txtQtdeC) + 1
                    Case "2": txtQtdeC = CInt(txtQtdeC) + 1
                    Case "X": txtQtdeX = CInt(txtQtdeX) + 1
                    Case "R": txtQtdeR = CInt(txtQtdeR) + 1
                End Select
            Else
                txtQtdeDEM = CInt(txtQtdeDEM) + 1
            End If
            adoQtde.MoveNext
        Loop
        txtQtdeTOTAL = CInt(txtQtdeG) + CInt(txtQtdeV) + CInt(txtQtdeC) + CInt(txtQtdeX) + CInt(txtQtdeR) + CInt(txtQtdeDEM)
        
        adoQtde.Close
    Else
        Grid.Height = 6915
        frmQtde.Visible = False
    End If
    
sair:
    Exit Sub
err1:
    If Err.Number = 3001 Then
       ' MsgBox "Dados inválidos para Filtragem!", vbCritical
    ElseIf Err.Number = 3021 Then
        MsgBox "Nenhum registro encontrado!", vbCritical
    ElseIf Err.Number <> 13 Then
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    End If
        W_LD_FILTRO = False
        Resume sair

End Sub



Private Sub cmdFiltrar_Click()
    
    On Error Resume Next
    
      'Dados Contas
    If (ADOREG.Recordset.Fields("M_TIPO") = "V" Or ADOREG.Recordset.Fields("M_TIPO") = "C") Or acessoTotal() Then
        If de.rscmdSqlVisAltContas.State = 1 Then de.rscmdSqlVisAltContas.Close
    Else
        If de.rscmdSqlVisAltContas2.State = 1 Then de.rscmdSqlVisAltContas2.Close
    End If
    
    FILTRAR W_INDEX

    Pause 0.5
                 
        'Dados Contas
        Lancamentos
  
  '  p_Pesq.Visible = False
    'If cbMostrar.ListIndex > 0 Then
    '    If w_reset_tipo Then
    '        cbMostrar.text = "TODOS"
    '    Else
    '        cmdMostrar_Click
    '    End If
    'End If
  
End Sub



Private Sub Total()
Dim ADO_TOTAL As ADODB.Recordset

On Error GoTo err1
    
    TXT_MAIS = 0
    TXT_MENOS = 0
    TXT_TOTAL = 0
    
    Set ADO_TOTAL = ADO_LANC.Recordset.Clone
    
    If Not ADO_TOTAL.EOF Then
        ADO_TOTAL.MoveFirst
        Do While Not ADO_TOTAL.EOF
            If ADO_TOTAL.Fields("VALOR") >= 0 And ADO_TOTAL.Fields("OP") = "+" Then
                TXT_MAIS = CDbl(TXT_MAIS) + ADO_TOTAL.Fields("VALOR")
            ElseIf ADO_TOTAL.Fields("VALOR") < 0 And ADO_TOTAL.Fields("OP") = "-" Then
                TXT_MENOS = CDbl(TXT_MENOS) + ADO_TOTAL.Fields("VALOR")
            End If
            ADO_TOTAL.MoveNext
        Loop
        
        TXT_TOTAL = CDbl(TXT_MAIS) + CDbl(TXT_MENOS)
    End If
    
    TXT_TOTAL = Format(TXT_TOTAL, "R$ 0.00")
    TXT_MAIS = Format(TXT_MAIS, "R$ #0.00")
    TXT_MENOS = Format(TXT_MENOS, "R$ #0.00")
    
    
    'muda cor do total
    If TXT_TOTAL < 0 Then
        TXT_TOTAL.ForeColor = vbRed
    Else
        TXT_TOTAL.ForeColor = vbWhite
    End If
    
    
    W_SALDO = de.cnc.Execute("Select F_SALDO_ANT FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & ADOREG.Recordset.Fields("M_F_COD") & "").Fields(0)
    'Saldo restante da ficha
    txt_SaldoAnt = IIf(IsNull(W_SALDO), 0, W_SALDO)
    If txt_SaldoAnt < 0 Then
        txt_SaldoAnt.ForeColor = vbRed
    Else
        txt_SaldoAnt.ForeColor = vbBlue
    End If
    txt_SaldoAnt = Format(txt_SaldoAnt, "R$ 0.00")


sair:
    Exit Sub
err1:
    'MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Sub Imprimir()
Dim SQL_Tripa As String
Dim SQL_TripaDet As String
Dim SQL_TripaFichas As String

Dim w_Lojas As String
Dim w_sqlLojasTripa As String

Dim w_Tipos As String
Dim w_sqlTiposTripa As String

Dim w_FirstLoja As Boolean
Dim w_FirstFicha As Boolean
Dim w_FirstTipo As Boolean
 
Dim w_tipoTripa(50) As Variant
Dim w_lojaTripa(100) As Variant
 
On Error GoTo err1
'On Error Resume Next
     
    FRM_IMP_F.Form_Activate
    If FRM_IMP_F.ckTodas.value = 0 Then
        FRM_IMP_F.TXT_LOGO = txtLogo
    Else
        FRM_IMP_F.TXT_LOGO = "%"
    End If
    
    FRM_IMP_F.TXT_MES = TXT_MES
    FRM_IMP_F.TXT_ANO = TXT_ANO
    
    If FRM_IMP_F.ck_Nome.value = 0 Then
        FRM_IMP_F.dbNome = TXT_FUNC
    Else
        FRM_IMP_F.dbNome = "%"
    End If
    FRM_IMP_F.CkFicha.Visible = True
    FRM_IMP_F.CkTripa.Visible = True
    

    FRM_IMP_F.txt_tipo = TXT_FTIPO
    
    FRM_IMP_F.Show 1
    
    If (FRM_IMP_F.txt_State = "F") Then 'Or (FRM_IMP_F.CkTripa.value = 1 And FRM_IMP_F.CkFicha.value = 1) Then
       MsgBox "Impressão Cancelada!", vbCritical
    Else
        
    'lojas
    w_FirstLoja = True
    For I = 0 To FRM_IMP_F.TXT_LOGO.ListCount - 1
        If FRM_IMP_F.TXT_LOGO.Selected(I) = True Then
            If w_FirstLoja Then
                w_Lojas = "'" & FRM_IMP_F.TXT_LOGO.list(I) & "'"
                w_sqlLojasTripa = " TAB_FUNCIONARIO.F_Cod_L = '" & FRM_IMP_F.TXT_LOGO.list(I) & "' "
            Else
                w_Lojas = w_Lojas & ",'" & FRM_IMP_F.TXT_LOGO.list(I) & "'"
                w_sqlLojasTripa = w_sqlLojasTripa & " OR TAB_FUNCIONARIO.F_Cod_L = '" & FRM_IMP_F.TXT_LOGO.list(I) & "' "
            End If
            w_lojaTripa(I) = FRM_IMP_F.TXT_LOGO.list(I)
            w_FirstLoja = False
        End If
    Next
    
    'tipos
    w_FirstTipo = True
    Dim w_tipo
    For J = 0 To FRM_IMP_F.txt_tipo.ListCount - 1
        If FRM_IMP_F.txt_tipo.Selected(J) = True Then
            Select Case FRM_IMP_F.txt_tipo.list(J)
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
                w_sqlTiposTripa = " TAB_FICHA_MENS.M_TIPO = '" & w_tipo & "' "
            Else
                w_Tipos = w_Tipos & ",'" & w_tipo & "'"
                w_sqlTiposTripa = w_sqlTiposTripa & " OR TAB_FICHA_MENS.M_TIPO = '" & w_tipo & "' "
            End If
            w_tipoTripa(J) = w_tipo
            w_FirstTipo = False
        End If
    Next
    
   
        If FRM_IMP_F.CkTripa.value = 1 Then
                
            
            If de.rscmdSqlTripa.State = 1 Then de.rscmdSqlTripa.Close
            
            'de.cmdRelFichaMensal_TRIPA FRM_IMP_F.TXT_MES, FRM_IMP_F.TXT_ANO, FRM_IMP_F.dbNome & "%", w_tipoTripa(0), w_tipoTripa(1) _
            '    , w_tipoTripa(2), w_tipoTripa(3), w_tipoTripa(4), w_tipoTripa(5), w_tipoTripa(6), w_tipoTripa(7), w_tipoTripa(8) _
            '    , w_tipoTripa(9), w_tipoTripa(10), w_lojaTripa(0), w_lojaTripa(1), w_lojaTripa(2), w_lojaTripa(3), w_lojaTripa(4) _
            '    , w_lojaTripa(5), w_lojaTripa(6), w_lojaTripa(7), w_lojaTripa(8), w_lojaTripa(9), w_lojaTripa(10), w_lojaTripa(11) _
            '    , w_lojaTripa(12), w_lojaTripa(13), w_lojaTripa(14), w_lojaTripa(15), w_lojaTripa(16), w_lojaTripa(17), w_lojaTripa(18) _
            '    , w_lojaTripa(19), w_lojaTripa(20), w_lojaTripa(21), w_lojaTripa(22), w_lojaTripa(23), w_lojaTripa(24), w_lojaTripa(25) _
            '    , w_lojaTripa(26), w_lojaTripa(27), w_lojaTripa(28), w_lojaTripa(29), w_lojaTripa(30), w_lojaTripa(31), w_lojaTripa(32) _
            '    , w_lojaTripa(33), w_lojaTripa(34), w_lojaTripa(35), w_lojaTripa(36), w_lojaTripa(37), w_lojaTripa(38), w_lojaTripa(39) _
            '    , w_lojaTripa(40), w_lojaTripa(41), w_lojaTripa(42), w_lojaTripa(43), w_lojaTripa(44), w_lojaTripa(45), w_lojaTripa(46) _
            '    , w_lojaTripa(47), w_lojaTripa(48), w_lojaTripa(49), w_lojaTripa(50), w_lojaTripa(51), w_lojaTripa(52), w_lojaTripa(53) _
            '    , w_lojaTripa(54), w_lojaTripa(55), w_lojaTripa(56), w_lojaTripa(57), w_lojaTripa(58), w_lojaTripa(59)
                        
            SQL_Tripa = "SELECT TAB_FICHA_MENS.M_NFICHA AS Ficha, TAB_FUNCIONARIO.F_NOME AS Nome," _
                & "Format('01/'+Mid(Str(TAB_FICHA_MENS.M_MES),2)+'/'+Mid(Str(TAB_FICHA_MENS.M_ANO),2),'DD/MM/YYYY') AS Data," _
                & "TAB_FUNCIONARIO.F_Cod_L AS Logo2, LOJB010.NUM as Logo, TAB_FICHA_MENS.M_TOTAL, Mid(TAB_FUNCIONARIO.F_COD_CENTRAL,3) AS COD_CENTRAL," _
                & " TAB_FICHA_MENS.M_TIPO AS TIPO, TAB_FUNCIONARIO.F_CX_QT_VND AS Cx_Qt_VND FROM TAB_FICHA_MENS, TAB_FUNCIONARIO INNER JOIN Lojb010 ON TAB_FUNCIONARIO.F_Cod_L = Lojb010.COD_LOJ " _
                & " WHERE (((TAB_FICHA_MENS.M_F_COD)=[TAB_FUNCIONARIO].[F_Codigo]) AND ((TAB_FICHA_MENS.M_MES)=" & FRM_IMP_F.TXT_MES & ") AND" _
                & " ((TAB_FICHA_MENS.M_ANO)=" & FRM_IMP_F.TXT_ANO & ") AND ((TAB_FUNCIONARIO.F_NOME) Like '" & FRM_IMP_F.dbNome & "' and TAB_FUNCIONARIO.F_NOME <> '10 - Func'" _
                & " AND TAB_FUNCIONARIO.F_NOME <> '99 - Presence') AND   (" _
                & w_sqlTiposTripa _
                & ") AND (" _
                & w_sqlLojasTripa _
                & ")) GROUP BY TAB_FICHA_MENS.M_NFICHA, TAB_FUNCIONARIO.F_NOME," _
                & " Format('01/'+Mid(Str(TAB_FICHA_MENS.M_MES),2)+'/'+Mid(Str(TAB_FICHA_MENS.M_ANO),2),'DD/MM/YYYY')," _
                & " TAB_FUNCIONARIO.F_Cod_L, TAB_FICHA_MENS.M_TOTAL, Mid(TAB_FUNCIONARIO.F_COD_CENTRAL,3), TAB_FICHA_MENS.M_TIPO," _
                & " TAB_FUNCIONARIO.F_CX_QT_VND, Len(TAB_FICHA_MENS.M_DT_ACF), LOJB010.NUM "
                
            If (FRM_IMP_F.ck_Nome.value) Then
                SQL_Tripa = SQL_Tripa & " HAVING ( (Len([TAB_FICHA_MENS].[M_DT_ACF]) IS NULL) OR (Len([TAB_FICHA_MENS].[M_DT_ACF])<5) ) "
            End If
            
            SQL_Tripa = SQL_Tripa & " ORDER BY LOJB010.NUM, TAB_FICHA_MENS.M_TIPO DESC , TAB_FUNCIONARIO.F_NOME; "
            
                
              'de.rscmdRelFichaMensal_TRIPA.Source = SQL_Tripa
                
            'TXT_OBS = SQL_Tripa
            de.rscmdSqlTripa.Open SQL_Tripa, , adOpenStatic, adLockOptimistic
            'de.rscmdRelFichaMensal_TRIPA.Open , , adOpenStatic, adLockOptimistic
            
            If Not de.rscmdSqlTripa.EOF Then
                If de.rscmdSqlTotalVND.State = 1 Then de.rscmdSqlTotalVND.Close
                
            
                w_DtI = CVDate("01/" & Format(FRM_IMP_F.TXT_MES, "00") & "/" & Format(FRM_IMP_F.TXT_ANO, "0000"))
                w_DtF = UltDiaMes(FRM_IMP_F.TXT_MES, FRM_IMP_F.TXT_ANO)
                de.cmdSqlTotalVND w_DtI, w_DtF, IIf(FRM_IMP_F.TXT_LOGO = "", "%", FRM_IMP_F.TXT_LOGO)
                
                
                'fichas
                w_FirstFicha = True
                
                de.rscmdSqlTripa.MoveFirst
                Do While Not de.rscmdSqlTripa.EOF
                    If w_FirstFicha Then
                        SQL_TripaFichas = " TAB_DESC_CALC.C_N_FICHA = " & de.rscmdSqlTripa.Fields("Ficha") & " "
                
                    Else
                        SQL_TripaFichas = SQL_TripaFichas & " OR TAB_DESC_CALC.C_N_FICHA = " & de.rscmdSqlTripa.Fields("Ficha") & " "
                    End If
                    w_FirstFicha = False
                    de.rscmdSqlTripa.MoveNext
                Loop

                de.rscmdSqlTripa.MoveFirst
                
                If de.rscmdSqlTripaDet.State = 1 Then de.rscmdSqlTripaDet.Close
                
                SQL_TripaDet = "SELECT TAB_DESC_CALC.C_N_FICHA AS Ficha, TAB_DESC_CALC.C_TP_CONTA, TAB_TP_CONTA.TP_DESC AS Conta," _
                    & "SUM(TAB_DESC_CALC.C_VALOR) AS Valor, TAB_DESC_CALC.C_TP_OP AS Op, TAB_TP_CONTA.TP_NIVEL FROM TAB_FICHA_MENS," _
                    & "TAB_DESC_CALC, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA " _
                    & "AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (" _
                    & SQL_TripaFichas

                
                If (de.rscmdSqlTripa.Fields("TIPO") = "V" Or de.rscmdSqlTripa.Fields("TIPO") = "C" Or de.rscmdSqlTripa.Fields("TIPO") = "X" Or de.rscmdSqlTripa.Fields("TIPO") = "2") Or acessoTotal() Then
                    SQL_TripaDet = SQL_TripaDet & ")"
                Else
                    SQL_TripaDet = SQL_TripaDet & ") AND ((TAB_DESC_CALC.C_TP_CONTA <> 20 and TAB_DESC_CALC.C_TP_CONTA <> 78 AND (TAB_FICHA_MENS.M_TIPO <> 'V' AND TAB_FICHA_MENS.M_TIPO <> 'C' AND TAB_FICHA_MENS.M_TIPO <> 'X' AND TAB_FICHA_MENS.M_TIPO <> '2')) OR (TAB_FICHA_MENS.M_TIPO = 'V' OR TAB_FICHA_MENS.M_TIPO = 'C' OR TAB_FICHA_MENS.M_TIPO = 'X' OR TAB_FICHA_MENS.M_TIPO = '2')) "
                End If
                
                    SQL_TripaDet = SQL_TripaDet & "GROUP BY TAB_DESC_CALC.C_N_FICHA, TAB_TP_CONTA.TP_DESC," _
                    & "TAB_DESC_CALC.C_TP_OP, TAB_DESC_CALC.C_TP_CONTA, TAB_TP_CONTA.TP_NIVEL ORDER BY TAB_DESC_CALC.C_N_FICHA," _
                    & "SUM(TAB_DESC_CALC.C_VALOR) DESC"
                
                de.rscmdSqlTripaDet.Open SQL_TripaDet, , adOpenStatic, adLockOptimistic
                
                
                'Set AdoItem1 = de.rscmdRelFichaMensal_TRIPA.Fields(8).value
                Set AdoItem1 = de.rscmdSqlTripaDet.Clone
                'Criar_RPT_TRIPA de.rscmdRelFichaMensal_TRIPA, AdoItem1
                PrintTripa de.rscmdSqlTripa, AdoItem1, (FRM_IMP_F.ck_Nome.value = 0 And FRM_IMP_F.ckTodas.value = 0)
                frmTripa.Show 1
            End If
        End If


        If FRM_IMP_F.CkFicha.value = 1 Then
        
        'Se quer varios emp. de uma vez então Filtra p/ não mostra os q/ tem acerto final
        w_ACF = ""
        If (FRM_IMP_F.dbNome = "%" Or FRM_IMP_F.ck_Nome = 1) Then
            'w_ACF = "and ((M_DT_ACF is null or M_DT_ACF ='') and M_BLOQ = 0)"
            w_ACF = "and ((M_DT_ACF is null or M_DT_ACF ='') OR M_BLOQ = 0)"
        End If

If (w_tipo = "V" Or w_tipo = "C" Or w_tipo = "X" Or w_tipo = "2") Or acessoTotal() Then
    sql_ficha = "SELECT TAB_FICHA_MENS.M_NFICHA, TAB_FICHA_MENS.M_LOGO, TAB_FICHA_MENS.M_MES, " _
                & "TAB_FICHA_MENS.M_ANO, TAB_FICHA_MENS.M_PG_VT, TAB_FICHA_MENS.M_NOTAS, TAB_FICHA_MENS.M_NOME AS NOME, TAB_FICHA_MENS.M_DT_ADM, TAB_FICHA_MENS.M_DT_REG, " _
                & "TAB_FICHA_MENS.M_DT_DEM, TAB_FICHA_MENS.M_FERIAS, TAB_FICHA_MENS.M_OBS, TAB_FICHA_MENS.M_ANOTACAO, " _
                & "CVDate(Format('01/'+Mid(Str([TAB_FICHA_MENS].[M_MES]),2)+'/'+Mid(Str([TAB_FICHA_MENS].[M_ANO]),2),'dd/mm/yyyy')) AS v_data, " _
                & "TAB_FICHA_MENS.M_TOTAL_MAIS, TAB_FICHA_MENS.M_TOTAL_MENOS, TAB_FICHA_MENS.M_TOTAL, TAB_DESC_CALC.C_DT, " _
                & "TAB_DESC_CALC.C_VALOR, TAB_DESC_CALC.C_TP_CONTA, TAB_DESC_CALC.C_TP_OP, TAB_DESC_CALC.C_DESC, TAB_DESC_CALC.C_VISTO, TAB_TP_CONTA.TP_DESC, " _
                & "TAB_TP_CONTA.TP_NIVEL as Ordem, TAB_FICHA_MENS.M_ACORDO, TAB_DESC_CALC.C_NCRED, " _
                & "TAB_FICHA_MENS.M_DT_ACF, TAB_FICHA_MENS.M_EMPRESTIMO_ANOT, TAB_FICHA_MENS.M_FERIAS_PG, TAB_FICHA_MENS.M_FERIAS_Ult_PG, " _
                & "TAB_FICHA_MENS.M_FERIAS_OK, TAB_FICHA_MENS.M_13_PG, TAB_FICHA_MENS.M_13_ULT_PG, TAB_FICHA_MENS.M_13_OBS, TAB_FICHA_MENS.M_13_OK, TAB_DESC_CALC.C_TP_CONTA as Conta , TAB_FICHA_MENS.M_TIPO as TIPO, " _
                & "TAB_FUNCIONARIO.F_VPISO AS PB, TAB_FUNCIONARIO.F_VPISO_R AS PL, TAB_FICHA_MENS.M_BLOQ, Lojb010.NUM, TAB_FUNCIONARIO.F_DT_NASC " _
                & "FROM (TAB_FUNCIONARIO INNER JOIN (TAB_TP_CONTA RIGHT JOIN (TAB_FICHA_MENS LEFT JOIN TAB_DESC_CALC ON TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA) ON " _
                & "TAB_TP_CONTA.TP_COD = TAB_DESC_CALC.C_TP_CONTA) ON TAB_FUNCIONARIO.F_Codigo = TAB_FICHA_MENS.M_F_COD) INNER JOIN Lojb010 ON TAB_FUNCIONARIO.F_Cod_L = Lojb010.COD_LOJ " _
                & " Where (( TAB_FICHA_MENS.M_LOGO IN (" & w_Lojas & ")) And ((TAB_FICHA_MENS.M_MES) = " & FRM_IMP_F.TXT_MES & ") " _
                & "And ((TAB_FICHA_MENS.M_ANO) = " & FRM_IMP_F.TXT_ANO & ") And NOT TAB_FICHA_MENS.M_NOME ='10 - Func' and not TAB_FICHA_MENS.M_NOME='99 - Presence' And ((TAB_FICHA_MENS.M_NOME) Like '" & IIf(FRM_IMP_F.dbNome = "%", "*", FRM_IMP_F.dbNome) & "') AND ((TAB_FICHA_MENS.M_TIPO) IN (" & w_Tipos & ")) " & IIf(w_ACF <> "", w_ACF, "") & ") " _
                & "ORDER BY TAB_FUNCIONARIO.F_Cod_L, TAB_FICHA_MENS.M_TIPO DESC, TAB_FUNCIONARIO.F_NOME, TAB_TP_CONTA.TP_NIVEL, TAB_DESC_CALC.C_TP_OP, TAB_FICHA_MENS.M_TIPO desc;"
               '& "ORDER BY TAB_FICHA_MENS.M_NFICHA, TAB_FICHA_MENS.M_NOME, TAB_TP_CONTA.TP_NIVEL, TAB_DESC_CALC.C_TP_OP;"
Else
        sql_ficha = "SELECT TAB_FICHA_MENS.M_NFICHA, TAB_FICHA_MENS.M_LOGO, TAB_FICHA_MENS.M_MES, " _
                & "TAB_FICHA_MENS.M_ANO, TAB_FICHA_MENS.M_NOTAS, TAB_FICHA_MENS.M_PG_VT, TAB_FICHA_MENS.M_NOME AS NOME, TAB_FICHA_MENS.M_DT_ADM, TAB_FICHA_MENS.M_DT_REG, " _
                & "TAB_FICHA_MENS.M_DT_DEM, TAB_FICHA_MENS.M_FERIAS, TAB_FICHA_MENS.M_OBS, TAB_FICHA_MENS.M_ANOTACAO, " _
                & "CVDate(Format('01/'+Mid(Str([TAB_FICHA_MENS].[M_MES]),2)+'/'+Mid(Str([TAB_FICHA_MENS].[M_ANO]),2),'dd/mm/yyyy')) AS v_data, " _
                & "TAB_FICHA_MENS.M_TOTAL_MAIS, TAB_FICHA_MENS.M_TOTAL_MENOS, TAB_FICHA_MENS.M_TOTAL, TAB_DESC_CALC.C_DT, " _
                & "TAB_DESC_CALC.C_VALOR, TAB_DESC_CALC.C_TP_CONTA, TAB_DESC_CALC.C_TP_OP, TAB_DESC_CALC.C_DESC, TAB_DESC_CALC.C_VISTO, TAB_TP_CONTA.TP_DESC, " _
                & "TAB_TP_CONTA.TP_NIVEL as Ordem, TAB_FICHA_MENS.M_ACORDO, TAB_DESC_CALC.C_NCRED, " _
                & "TAB_FICHA_MENS.M_DT_ACF, TAB_FICHA_MENS.M_EMPRESTIMO_ANOT, TAB_FICHA_MENS.M_FERIAS_PG, TAB_FICHA_MENS.M_FERIAS_Ult_PG, " _
                & "TAB_FICHA_MENS.M_FERIAS_OK, TAB_FICHA_MENS.M_13_PG, TAB_FICHA_MENS.M_13_ULT_PG, TAB_FICHA_MENS.M_13_OBS, TAB_FICHA_MENS.M_13_OK, TAB_DESC_CALC.C_TP_CONTA as Conta , TAB_FICHA_MENS.M_TIPO as TIPO, " _
                & "TAB_FUNCIONARIO.F_VPISO AS PB, TAB_FUNCIONARIO.F_VPISO_R AS PL, TAB_FICHA_MENS.M_BLOQ, Lojb010.NUM, TAB_FUNCIONARIO.F_DT_NASC " _
                & "FROM (TAB_FUNCIONARIO INNER JOIN (TAB_TP_CONTA RIGHT JOIN (TAB_FICHA_MENS LEFT JOIN TAB_DESC_CALC ON TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA) ON " _
                & "TAB_TP_CONTA.TP_COD = TAB_DESC_CALC.C_TP_CONTA) ON TAB_FUNCIONARIO.F_Codigo = TAB_FICHA_MENS.M_F_COD) INNER JOIN Lojb010 ON TAB_FUNCIONARIO.F_Cod_L = Lojb010.COD_LOJ " _
                & " Where (( ((TAB_DESC_CALC.C_TP_CONTA <> 20 and TAB_DESC_CALC.C_TP_CONTA <> 78 AND (TAB_FICHA_MENS.M_TIPO <> 'V' AND TAB_FICHA_MENS.M_TIPO <> 'C' AND TAB_FICHA_MENS.M_TIPO <> 'X' AND TAB_FICHA_MENS.M_TIPO <> '2')) " _
                & " OR (TAB_FICHA_MENS.M_TIPO = 'V' OR TAB_FICHA_MENS.M_TIPO = 'C' OR TAB_FICHA_MENS.M_TIPO = 'X' OR TAB_FICHA_MENS.M_TIPO = '2') ) AND  TAB_FICHA_MENS.M_LOGO IN (" & w_Lojas & ")) And ((TAB_FICHA_MENS.M_MES) = " & FRM_IMP_F.TXT_MES & ") " _
                & "And ((TAB_FICHA_MENS.M_ANO) = " & FRM_IMP_F.TXT_ANO & ") And NOT TAB_FICHA_MENS.M_NOME ='10 - Func' and not TAB_FICHA_MENS.M_NOME='99 - Presence' And ((TAB_FICHA_MENS.M_NOME) Like '" & IIf(FRM_IMP_F.dbNome = "%", "*", FRM_IMP_F.dbNome) & "') AND ((TAB_FICHA_MENS.M_TIPO) IN (" & w_Tipos & ")) " & IIf(w_ACF <> "", w_ACF, "") & ") " _
                & "ORDER BY TAB_FUNCIONARIO.F_Cod_L, TAB_FICHA_MENS.M_TIPO DESC, TAB_FUNCIONARIO.F_NOME, TAB_TP_CONTA.TP_NIVEL, TAB_DESC_CALC.C_TP_OP, TAB_FICHA_MENS.M_TIPO desc;"
End If
        'TXT_OBS.Text = SQL_Ficha
        
        If de.rsTab_Config.State = 0 Then de.Tab_Config
        de.rsTab_Config.Fields("SQL_RPT") = sql_ficha
        de.rsTab_Config.UpdateBatch adAffectCurrent

 '       de.cnc.BeginTrans
        de.rsTab_Config.Fields("SQL_RPT") = sql_ficha
        de.cnc.BeginTrans
        de.rsTab_Config.UpdateBatch adAffectCurrent
        de.cnc.CommitTrans
        
Dim w_Access As Access.Application
Set w_Access = New Access.Application
    
    If acessoTotal() Then
      
        If MsgBox("Deseja exibir as ANOTAÇÕES EXTRAS da ficha?", vbYesNo, "Impressão de Ficha") = vbYes Then
            w_Access.OpenCurrentDatabase Left(strDirBase, Len(strDirBase) - 9) & "rptNotas.mdb", False
        Else
            w_Access.OpenCurrentDatabase Left(strDirBase, Len(strDirBase) - 9) & "rpt.mdb", False
        End If
        
    Else
    
            w_Access.OpenCurrentDatabase Left(strDirBase, Len(strDirBase) - 9) & "rpt.mdb", False
    End If
       w_Access.Reports(REL_FICHA_MENS).Printer.ColorMode = acPRCMColor 'imprimir colorido, sem marcar aquele "IMPRIMIR EM ESCALA DE CINZA"
       w_Access.DoCmd.OpenReport ReportName:="REL_FICHA_MENS", View:=Access.acViewPreview, WindowMode:=Access.acWindowNormal
       
        'w_Access.Visible = True
        'w_Access.DoCmd.RunCommand acCmdDocMaximize
        'rptFichaMensal.Show
    End If
    
    
    End If
    
sair:
    Exit Sub
err1:
    w_Access.CloseCurrentDatabase
    Set w_Access = Nothing
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume sair
    
End Sub

Private Sub Refresh_dados()
On Error Resume Next
    W_FILTRO = ADOREG.Recordset.Filter
    pos = ADOREG.Recordset.Fields("m_nficha")


'    ADOREG.Recordset.CancelBatch adAffectCurrent
'    de.rscmdSqlVisAltFichas.Resync
    
On Error GoTo err1
 '   If de.rscmdSqlVisAltFichas.State = 1 Then de.rscmdSqlVisAltFichas.Close
 '   de.cmdSqlVisAltFichas txt_PMes, txt_PAno
 '   Set ADOREG.Recordset = de.rscmdSqlVisAltFichas.Clone
'    Set ADO_LANC.Recordset = ADOREG.Recordset.Fields("cmdSqlVisAltContas").UnderlyingValue
    
    
    If W_FILTRO <> "0" Then ADOREG.Recordset.Filter = W_FILTRO
    ADOREG.Recordset.MoveFirst
    If pos <> Empty Then
        Do While Not ADOREG.Recordset.EOF
          If ADOREG.Recordset.Fields("m_nficha") = pos Then Exit Do
          ADOREG.Recordset.Find "m_nficha = " & pos & ""
        Loop
    End If
       '*** DESABILITA O EDITAR ****
   If ADOREG.Recordset.Fields("M_BLOQ") = True Then
        BarraF.Buttons("editar").Enabled = False
   Else
        BarraF.Buttons("editar").Enabled = True
   End If
    
sair:


    Exit Sub
err1:
    If Err.Number = 3021 Then
        Form_Load
    Else
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    End If
    Resume sair
End Sub

'***MENU ***

Private Sub mnuAcessoTotal_Click()
    If acessoTotal() Then
        w_usuario = "USER"
        Lancamentos
        cmdEsconder.BackColor = vbRed
        txt_notas.Visible = False
        lblNotas.Visible = False
        BarraF.Buttons("desbloquear").Enabled = False
        cmdDesbloquear.Visible = False
    ElseIf w_usuario = "USER" Then
        w_usuario = w_usuario2
        Lancamentos
        cmdEsconder.BackColor = &H8000000F
        txt_notas.Visible = True
        lblNotas.Visible = True
        If (UCase(w_usuario) = UCase(NomeMestre) Or UCase(w_usuario) = UCase(NomeMestre2) Or UCase(w_usuario) = UCase(NomeMestre3)) Then
            BarraF.Buttons("desbloquear").Enabled = True
            cmdDesbloquear.Visible = True
        End If
    End If
    Grid.SetFocus

End Sub


Private Sub mnuRemT_Click()
On erro GoTo err1

If BarraF.Buttons("editar").Enabled = True Then
    
    'If Not isMesValido(txt_F_COD, TXT_MES, TXT_ANO) Then 'Verifica se é mês atual ou passado
        'If MsgBox("Você está alterando uma ficha que NÃO É DO MÊS ATUAL. Deseja continuar mesmo assim?", vbYesNo, "Alteração de ficha") = vbNo Then
        '    Exit Sub
        'End If
        If ADOREG.Recordset.Fields("M_BLOQ") Then
            MsgBox "Esta ficha não é do mês atual e está BLOQUEADA!", vbCritical
            Exit Sub
        End If
    'End If


    'Atualiza Visto
    w_cod = flexGRID_L.TextMatrix(flexGRID_L.RowSel, 7)
    W_NFICHA = TXT_NFICHA
    W_F_COD = txt_F_COD
    
    
    'w_cod = ADO_LANC.Recordset.Fields("C_Codigo")
    'W_NFICHA = adoReg.Recordset.Fields("M_NFICHA")
    'W_F_COD = adoReg.Recordset.Fields("M_F_COD")
        
    '*** SE EXISTIR ALGUMA CONTA TIPO   FERIAS 24
    If de.cnc.Execute("SELECT C_TP_CONTA FROM TAB_DESC_CALC WHERE C_N_FICHA = " & W_NFICHA & " AND C_TP_CONTA = 24").RecordCount > 0 Then
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS_OK = 0  WHERE (F_CODIGO = " & W_F_COD & " )", w_reg
            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_FERIAS_OK = 0  WHERE (M_NFICHA = " & W_NFICHA & " )", w_reg
            CK_FERIAS.value = 0
    End If
    
    '*** SE EXISTIR ALGUMA CONTA TIPO   13º  32
    If de.cnc.Execute("SELECT C_TP_CONTA FROM TAB_DESC_CALC WHERE C_N_FICHA = " & W_NFICHA & " AND C_TP_CONTA = 32").RecordCount > 0 Then
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_13_OK = 0  WHERE (F_CODIGO = " & W_F_COD & " )", w_reg
            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_13_OK = 0  WHERE (M_NFICHA = " & W_NFICHA & " )", w_reg
            CK_13.value = 0
    End If

    'Atualiza Visto em todos os lançamentos da Ficha Corrente ***    REMOVER
    de.cnc.Execute "Update TAB_DESC_CALC Set C_VISTO = 0 Where (C_N_FICHA = " & W_NFICHA & ")"
    
    'Atualiza Visto em todos os lançamentos fixos do funcionário **REMOVER
    'de.cnc.Execute "Update TAB_DESC_CALC_FIXO Set CF_VISTO = 0 Where (CF_EMP_COD = " & W_F_COD & ")"
    
    
    Lancamentos
    'GRID_L.ReBind
    'GRID_L.Refresh
    
    flexGRID_L.Refresh
    
End If

sair:
    Exit Sub
err1:
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub mnuVist_Click()
On erro GoTo err1

If BarraF.Buttons("editar").Enabled = True Then

    'If Not isMesValido(txt_F_COD, TXT_MES, TXT_ANO) Then 'Verifica se é mês atual ou passado
        'If MsgBox("Você está alterando uma ficha que NÃO É DO MÊS ATUAL. Deseja continuar mesmo assim?", vbYesNo, "Alteração de ficha") = vbNo Then
        '    Exit Sub
        'End If
        If ADOREG.Recordset.Fields("M_BLOQ") Then
            MsgBox "Esta ficha não é do mês atual e está BLOQUEADA!", vbCritical
            Exit Sub
        End If
    'End If

    'Atualiza Visto
    w_cod = flexGRID_L.TextMatrix(flexGRID_L.RowSel, 7)
    W_NFICHA = TXT_NFICHA
    W_F_COD = txt_F_COD
    
    'w_cod = ADO_LANC.Recordset.Fields("C_Codigo")
    'W_NFICHA = adoReg.Recordset.Fields("M_NFICHA")
    'W_F_COD = adoReg.Recordset.Fields("M_F_COD")
    
    '*** SE EXISTIR ALGUMA CONTA TIPO   FERIAS 24
    If de.cnc.Execute("SELECT C_TP_CONTA FROM TAB_DESC_CALC WHERE C_N_FICHA = " & W_NFICHA & " AND C_TP_CONTA = 24").RecordCount > 0 Then
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS_OK = -1  WHERE (F_CODIGO = " & W_F_COD & " )", w_reg
            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_FERIAS_OK = -1  WHERE (M_NFICHA = " & W_NFICHA & " )", w_reg
            CK_FERIAS.value = 1
    End If
    
    '*** SE EXISTIR ALGUMA CONTA TIPO   13º  32
    If de.cnc.Execute("SELECT C_TP_CONTA FROM TAB_DESC_CALC WHERE C_N_FICHA = " & W_NFICHA & " AND C_TP_CONTA = 32").RecordCount > 0 Then
            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_13_OK = -1  WHERE (F_CODIGO = " & W_F_COD & " )", w_reg
            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_13_OK = -1  WHERE (M_NFICHA = " & W_NFICHA & " )", w_reg
            CK_13.value = 1
    End If


    'Atualiza Visto em todos os lançamentos da Ficha Corrente ***   VISTAR
    de.cnc.Execute "Update TAB_DESC_CALC Set C_VISTO = -1 Where (C_N_FICHA = " & W_NFICHA & ")"
    
    'Atualiza Visto em todos os lançamentos fixos do funcionário
    'de.cnc.Execute "Update TAB_DESC_CALC_FIXO Set CF_VISTO = -1 Where (CF_EMP_COD = " & W_F_COD & ")"
    
    
    Lancamentos
    'GRID_L.ReBind
    'GRID_L.Refresh
    flexGRID_L.Refresh

End If
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub mnuVis_Click()
On erro GoTo err1
    
 If BarraF.Buttons("editar").Enabled = True Then
 
    'If Not isMesValido(txt_F_COD, TXT_MES, TXT_ANO) Then 'Verifica se é mês atual ou passado
        'If MsgBox("Você está alterando uma ficha que NÃO É DO MÊS ATUAL. Deseja continuar mesmo assim?", vbYesNo, "Alteração de ficha") = vbNo Then
        '    Exit Sub
        'End If
        If ADOREG.Recordset.Fields("M_BLOQ") Then
            MsgBox "Esta ficha não é do mês atual e está BLOQUEADA!", vbCritical
            Exit Sub
        End If
    'End If
 
    'Atualiza Visto
    w_cod = flexGRID_L.TextMatrix(flexGRID_L.RowSel, 7)
    W_NFICHA = TXT_NFICHA
    W_F_COD = txt_F_COD
    
    'w_cod = ADO_LANC.Recordset.Fields("C_Codigo")
    'W_NFICHA = ADOREG.Recordset.Fields("M_NFICHA")
    'W_F_COD = ADOREG.Recordset.Fields("M_F_COD")
        
    '*** ATUALIZA TAB_FUNCIONARIO O CAMPO OK   SE   FOR   FERIAS OU 13ºSALARIO
    If flexGRID_L.TextMatrix(flexGRID_L.RowSel, 1) = "24" Then      'FERIAS
        de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS_OK = -1  WHERE (F_CODIGO = " & W_F_COD & " )", w_reg
        de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_FERIAS_OK = -1  WHERE (M_NFICHA = " & W_NFICHA & " )", w_reg
        CK_FERIAS.value = 1
    ElseIf flexGRID_L.TextMatrix(flexGRID_L.RowSel, 1) = "32" Then   '*** 13º
        de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_13_OK = -1  WHERE (F_CODIGO = " & W_F_COD & " )", w_reg
        de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_13_OK = -1  WHERE (M_NFICHA = " & W_NFICHA & " )", w_reg
        CK_13.value = 1
    End If
    'de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS_PG = NULL  WHERE (F_Codigo = " & TXT_FUNC.BoundText & " )", w_reg
 
    
    de.cnc.Execute "Update TAB_DESC_CALC Set C_VISTO = -1 Where (C_CODIGO = " & w_cod & ")"
    
    'Quando for fixo, coloca o visto da tabela de fixo
    'If flexGRID_L.TextMatrix(flexGRID_L.RowSel, 7) > 0 Then
    '   de.cnc.Execute "Update TAB_DESC_CALC_FIXO Set CF_VISTO = -1 Where (CF_CODIGO = " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 7) & ")"
    'End If
    
    Lancamentos
    'GRID_L.ReBind
    'GRID_L.Refresh
    flexGRID_L.Refresh
    ADO_LANC.Recordset.Find "C_Codigo = " & w_cod
    
 End If
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub mnuRem_Click()
On erro GoTo err1


If BarraF.Buttons("editar").Enabled = True Then

    'If Not isMesValido(txt_F_COD, TXT_MES, TXT_ANO) Then 'Verifica se é mês atual ou passado
        'If MsgBox("Você está alterando uma ficha que NÃO É DO MÊS ATUAL. Deseja continuar mesmo assim?", vbYesNo, "Alteração de ficha") = vbNo Then
        '    Exit Sub
        'End If
        If ADOREG.Recordset.Fields("M_BLOQ") Then
            MsgBox "Esta ficha não é do mês atual e está BLOQUEADA!", vbCritical
            Exit Sub
        End If
    'End If
    
    'Atualiza Visto
    w_cod = flexGRID_L.TextMatrix(flexGRID_L.RowSel, 7)
    W_NFICHA = TXT_NFICHA
    W_F_COD = txt_F_COD
    
    'w_cod = ADO_LANC.Recordset.Fields("C_Codigo")
    'W_NFICHA = ADOREG.Recordset.Fields("M_NFICHA")
    'W_F_COD = ADOREG.Recordset.Fields("M_F_COD")
    
    
    '*** ATUALIZA TAB_FUNCIONARIO O CAMPO OK   SE   FOR   FERIAS OU 13ºSALARIO
    If flexGRID_L.TextMatrix(flexGRID_L.RowSel, 1) = "24" Then      'FERIAS
        de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS_OK = 0  WHERE (F_CODIGO = " & W_F_COD & " )", w_reg
        de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_FERIAS_OK = 0  WHERE (M_NFICHA = " & W_NFICHA & " )", w_reg
        CK_FERIAS.value = 0
    ElseIf flexGRID_L.TextMatrix(flexGRID_L.RowSel, 1) = "32" Then   '*** 13º
        de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_13_OK = 0  WHERE (F_CODIGO = " & W_F_COD & " )", w_reg
        de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_13_OK = 0  WHERE (M_NFICHA = " & W_NFICHA & " )", w_reg
        CK_13.value = 0
    End If
    
    
    
    de.cnc.Execute "Update TAB_DESC_CALC Set C_VISTO = 0 Where (C_CODIGO = " & w_cod & ")"
    
    'Quando for fixo, tira o visto da tabela de fixo
    'If flexGRID_L.TextMatrix(flexGRID_L.RowSel, 7) > 0 Then
    '   de.cnc.Execute "Update TAB_DESC_CALC_FIXO Set CF_VISTO = 0 Where (CF_CODIGO = " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 7) & ")"
    'End If
    
    Lancamentos
    'GRID_L.ReBind
    'GRID_L.Refresh
    flexGRID_L.Refresh
    
    ADO_LANC.Recordset.Find "C_Codigo = " & w_cod
    

End If
    
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair

End Sub




Private Sub ExibirCelula()
    Static OK As Boolean
    
    wTxtOld = ""
    
    ' Se for celula fixa , sair
    If flexGRID_L.Col <= flexGRID_L.FixedCols - 1 Or flexGRID_L.Row <= flexGRID_L.FixedRows - 1 Then
        Exit Sub
    End If
     
    If (flexGRID_L.ColSel <= 2) Or (flexGRID_L.ColSel > 4) Then
        Exit Sub
    End If
    
    If OK Then Exit Sub
    OK = True
    
    OcultarControles
    
    LastRow = flexGRID_L.Row
    LastCol = flexGRID_L.Col
    
    'Nova Celula
    With flexGRID_L
        If .TextMatrix(LastRow, 0) = NovaLinha Then
            .Rows = .Rows + 1
            .TextMatrix(LastRow, 0) = LastRow
            .TextMatrix(.Rows - 1, 0) = NovaLinha
       End If
    End With
    
    Select Case LastCol
    Case Else
        Text1.Move flexGRID_L.CellLeft - Screen.TwipsPerPixelX, flexGRID_L.CellTop + 7045 - Screen.TwipsPerPixelY, flexGRID_L.CellWidth + Screen.TwipsPerPixelX * 2, flexGRID_L.CellHeight + Screen.TwipsPerPixelY * 2
        Text1.text = flexGRID_L.text
        If Len(flexGRID_L.text) = 0 Then
            If LastRow > 1 Then
                Text1.text = flexGRID_L.TextMatrix(LastRow - 1, LastCol)
            End If
        End If
        Text1.Visible = True
        If Text1.Visible Then
            Text1.ZOrder
            Text1.SetFocus
        End If
    End Select
    
    ControlVisible = True
    OK = False
    
    wTxtOld = Text1.text

End Sub
Private Sub ProximaCelula()
    If flexGRID_L.Col < flexGRID_L.Cols - 1 Then
        flexGRID_L.Col = flexGRID_L.Col + 1
    Else
        flexGRID_L.Col = 1
        If flexGRID_L.Row < flexGRID_L.Rows - 1 Then
            flexGRID_L.Row = flexGRID_L.Row + 1
        End If
    End If
End Sub
Private Sub AtribuiValorCelula()
    Dim texto As String
    Dim Op As String
    texto = Text1.text
    
    If texto <> flexGRID_L.TextMatrix(flexGRID_L.RowSel, flexGRID_L.ColSel) Then 'Se houve alteração
    
        Op = flexGRID_L.TextMatrix(flexGRID_L.RowSel, 5) 'op
        
        If flexGRID_L.ColSel = 4 Then 'Se Valor (e não digitou número)
            If Not (IsNumeric(texto)) Then
                MsgBox "Digite algum valor válido ou [ESC] para CANCELAR!", vbCritical, "Valor inválido"
                Exit Sub
            End If
            
            If (texto < 0 And (Op = "+" Or Op = "=")) Or (texto > 0 And Op = "-") Then texto = texto * -1
        End If
        
        If flexGRID_L.ColSel = 5 Then 'Se OP (e não digitou sinal)
            If Not (texto = "+" Or texto = "=" Or texto = "-") Then
                MsgBox "Digite algum sinal de -, + ou =!", vbCritical, "Sinal inválido"
                Exit Sub
            End If
            
        End If
        
       If flexGRID_L.ColSel = 0 Then 'Se Data (e não digitou data)
            If Not (IsDate(texto)) Then
                MsgBox "Digite alguma data válida ou [ESC] para CANCELAR!", vbCritical, "Data inválida"
                Exit Sub
            End If
        End If
        
            
        If (MsgBox("Deseja salvar as alterações?", vbYesNo, "Gravar alterações") = vbYes) Then

        flexGRID_L.TextMatrix(LastRow, LastCol) = texto
        flexGRID_L.CellForeColor = vbBlue
        
            If flexGRID_L.ColSel = 0 Then 'Data
                de.cnc.Execute ("UPDATE TAB_DESC_CALC set C_DT = '" & CDate(texto) & "' WHERE C_CODIGO = " & Str(flexGRID_L.TextMatrix(flexGRID_L.RowSel, 7)))
                de.cmdIncluirLog Date, Time, w_usuario, "EDITAR", "LANÇAMENTOS", "FICHA: " & TXT_NFICHA & " | DATA: " & texto & " | VALOR: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 4) & " | CONTA COD: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 1) & " | CONTA E DESCRICAO: " & texto & " | OP: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 5) & "   >>> DATA ANTERIOR: " & wTxtOld
            ElseIf flexGRID_L.ColSel = 3 Then 'Descricao conta
                de.cnc.Execute ("UPDATE TAB_DESC_CALC set C_DESC = '" & texto & "' WHERE C_CODIGO = " & Str(flexGRID_L.TextMatrix(flexGRID_L.RowSel, 7)))
                de.cmdIncluirLog Date, Time, w_usuario, "EDITAR", "LANÇAMENTOS", "FICHA: " & TXT_NFICHA & " | DATA: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 0) & " | VALOR: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 4) & " | CONTA COD: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 1) & " | DESCRICAO: " & texto & " | OP: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 5) & "   >>> DESCRIÇÃO ANTERIOR: " & wTxtOld
            ElseIf flexGRID_L.ColSel = 4 Then 'Valor
                de.cnc.Execute ("UPDATE TAB_DESC_CALC set C_VALOR = " & Str(texto) & " WHERE C_CODIGO = " & Str(flexGRID_L.TextMatrix(flexGRID_L.RowSel, 7)))
                de.cmdIncluirLog Date, Time, w_usuario, "EDITAR", "LANÇAMENTOS", "FICHA: " & TXT_NFICHA & " | DATA: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 0) & " | VALOR: " & Str(texto) & " | CONTA COD: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 1) & " | DESCRICAO: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 3) & " | OP: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 5) & "   >>> VALOR ANTERIOR: " & wTxtOld
            ElseIf flexGRID_L.ColSel = 5 Then 'OP
                de.cnc.Execute ("UPDATE TAB_DESC_CALC set C_TP_OP = '" & texto & "' WHERE C_CODIGO = " & Str(flexGRID_L.TextMatrix(flexGRID_L.RowSel, 7)))
                de.cmdIncluirLog Date, Time, w_usuario, "EDITAR", "LANÇAMENTOS", "FICHA: " & TXT_NFICHA & " | DATA: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 0) & " | VALOR: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 4) & " | CONTA COD: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 1) & " | DESCRICAO: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 3) & " | OP: " & texto & "   >>> OP ANTERIOR: " & wTxtOld
            End If
            
        End If
    End If
    
    OcultarControles
    ControlVisible = False
    
    Lancamentos
    
    
End Sub
Private Sub OcultarControles()
    ' Ocultar o controle textbox
    Text1.text = ""
    Text1.Visible = False
End Sub

Private Sub txtLogo2_Click(Area As Integer)
    txtLogo.BoundText = txtLogo2.BoundText
End Sub

Private Sub txtQtdeLimiteV_DblClick()
    Sendkeys "{home}+{end}"
End Sub


Private Sub txtQtdeLimiteV_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    If CInt(txtQtdeLimiteV) <> CInt(wTxtOld) Then
       If (MsgBox("Deseja salvar as alterações?", vbYesNo, "Gravar alterações") = vbYes) Then
           de.cnc.Execute ("UPDATE lojb010 SET QtdeLimiteVend = " & txtQtdeLimiteV & " WHERE COD_LOJ = '" & ADOREG.Recordset.Fields("M_LOGO") & "'")
           wTxtOld = txtQtdeLimiteV
           Sendkeys "{tab}"
       Else
           Sendkeys "{tab}"
       End If
    Else
        Sendkeys "{tab}"
    End If
  ElseIf KeyCode = vbKeyEscape Then
    txtQtdeLimiteV = wTxtOld
    Sendkeys "{tab}"
  End If
End Sub

Private Sub txtQtdeLimiteV_LostFocus()
    If CInt(txtQtdeLimiteV) <> CInt(wTxtOld) Then
       If (MsgBox("Deseja salvar as alterações?", vbYesNo, "Gravar alterações") = vbYes) Then
           de.cnc.Execute ("UPDATE lojb010 SET QtdeLimiteVend = " & txtQtdeLimiteV & " WHERE COD_LOJ = '" & ADOREG.Recordset.Fields("M_LOGO") & "'")
       Else
           txtQtdeLimiteV = wTxtOld
       End If
    End If
End Sub

