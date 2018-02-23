VERSION 5.00
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frmMenu 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   " M E N U"
   ClientHeight    =   8745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleMode       =   0  'User
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   " RELATÓRIOS "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   6240
      TabIndex        =   15
      Top             =   120
      Width           =   3135
      Begin Skin_Button.ctr_Button ctr_Button10 
         Height          =   405
         Left            =   240
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   360
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Qtde de Emp. por Logo"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button ctr_Button11 
         Height          =   405
         Left            =   240
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   960
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Comparativo de Alterações"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button ctr_Button12 
         Height          =   405
         Left            =   240
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Resumo Logos - Analítico"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button ctr_Button13 
         Height          =   405
         Left            =   240
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2040
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Resumo Logos - Sintético"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button ctr_Button14 
         Height          =   405
         Left            =   240
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Resumo T.P"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button ctr_Button15 
         Height          =   405
         Left            =   240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3120
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Empréstimo (pagos)"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":008C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button ctr_Button16 
         Height          =   405
         Left            =   240
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Empréstimo - Análise"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":00A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button ctr_Button17 
         Height          =   405
         Left            =   240
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   5760
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Vendas com Prêmio"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":00C4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button ctr_Button18 
         Height          =   405
         Left            =   240
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   4200
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Salário dos Cxs"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":00E0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button cmdCodMwts 
         Height          =   405
         Left            =   240
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   5280
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Código Mwts"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":00FC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button cmdRelFixosSaldos 
         Height          =   405
         Left            =   240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   6960
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Programados e Saldos Negativos"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":0118
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button cmdRelSalG 
         Height          =   405
         Left            =   240
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   4680
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Salários Gerentes"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":0134
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button cmdRelSalarios 
         Height          =   405
         Left            =   240
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   6360
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Conferência Comis/Prêmio/Piso"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":0150
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button cmdSalarios 
         Height          =   405
         Left            =   240
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   8040
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Salários Brutos"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":016C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button cmdSemRegistro 
         Height          =   405
         Left            =   240
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   7440
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Sem Registro"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":0188
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Caption         =   " ROTINAS "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Width           =   6015
      Begin Skin_Button.ctr_Button ctr_Button4 
         Height          =   525
         Left            =   2160
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "Lançamentos"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":01A4
         PICN            =   "frmMenu.frx":01C0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button ctr_Button5 
         Height          =   525
         Left            =   2160
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "Salário Família"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":0A9A
         PICN            =   "frmMenu.frx":0AB6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button ctr_Button6 
         Height          =   525
         Left            =   4080
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "Acesso Especial"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":26D8
         PICN            =   "frmMenu.frx":26F4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button ctr_Button7 
         Height          =   525
         Left            =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "Gerar Fichas"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":2B46
         PICN            =   "frmMenu.frx":2B62
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button ctr_Button8 
         Height          =   525
         Left            =   4080
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   360
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "Vistar Contas"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":2E7C
         PICN            =   "frmMenu.frx":2E98
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button ctr_Button9 
         Height          =   525
         Left            =   240
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "BACKUP"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":32EA
         PICN            =   "frmMenu.frx":3306
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button ctr_Button19 
         Height          =   525
         Left            =   2160
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "LOG"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":3620
         PICN            =   "frmMenu.frx":363C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button ctr_Button20 
         Height          =   525
         Left            =   240
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "Empréstimos"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenu.frx":3A8E
         PICN            =   "frmMenu.frx":3AAA
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   " CADASTROS "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6015
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   " GRIDs "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   3480
         TabIndex        =   32
         Top             =   360
         Width           =   2295
         Begin Skin_Button.ctr_Button cmdGridGerente 
            Height          =   525
            Left            =   240
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   926
            BTYPE           =   3
            TX              =   "GERENTES"
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
            COLTYPE         =   4
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmMenu.frx":3DC4
            PICN            =   "frmMenu.frx":3DE0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Skin_Button.ctr_Button cmdGridCaixa 
            Height          =   525
            Left            =   240
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   960
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   926
            BTYPE           =   3
            TX              =   "CAIXAS"
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
            COLTYPE         =   4
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmMenu.frx":6592
            PICN            =   "frmMenu.frx":65AE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Skin_Button.ctr_Button cmdGridVendedor 
            Height          =   525
            Left            =   240
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   360
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   926
            BTYPE           =   3
            TX              =   "VENDEDORES"
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
            COLTYPE         =   4
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmMenu.frx":8D60
            PICN            =   "frmMenu.frx":8D7C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Skin_Button.ctr_Button cmdGridLoja 
            Height          =   525
            Left            =   240
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   926
            BTYPE           =   3
            TX              =   "LOJAS"
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
            COLTYPE         =   4
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmMenu.frx":B52E
            PICN            =   "frmMenu.frx":B54A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Skin_Button.ctr_Button cmdGridMeta 
            Height          =   525
            Left            =   240
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   2880
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   926
            BTYPE           =   3
            TX              =   "METAS"
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
            COLTYPE         =   4
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmMenu.frx":DCFC
            PICN            =   "frmMenu.frx":DD18
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Caption         =   " Tipo de Conta "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   3015
         Begin Skin_Button.ctr_Button ctr_Button2 
            Height          =   525
            Left            =   240
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   360
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   926
            BTYPE           =   3
            TX              =   "NOVO Tipo de Conta"
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
            COLTYPE         =   4
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmMenu.frx":104CA
            PICN            =   "frmMenu.frx":104E6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Skin_Button.ctr_Button ctr_Button3 
            Height          =   525
            Left            =   240
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1080
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   926
            BTYPE           =   3
            TX              =   "Alterar Tipo de Conta"
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
            COLTYPE         =   4
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmMenu.frx":121F0
            PICN            =   "frmMenu.frx":1220C
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   " Funcionários "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3015
         Begin Skin_Button.ctr_Button cmdAddLanç_SalF 
            Height          =   525
            Left            =   240
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   360
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   926
            BTYPE           =   3
            TX              =   "NOVO Funcionário"
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
            COLTYPE         =   4
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmMenu.frx":12526
            PICN            =   "frmMenu.frx":12542
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Skin_Button.ctr_Button ctr_Button1 
            Height          =   525
            Left            =   240
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1080
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   926
            BTYPE           =   3
            TX              =   "Funcionário Existente"
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
            COLTYPE         =   4
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmMenu.frx":1424C
            PICN            =   "frmMenu.frx":14268
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
   End
   Begin Skin_Button.ctr_Button btnFichasMensais 
      Height          =   645
      Left            =   2640
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   1138
      BTYPE           =   3
      TX              =   " FICHAS MENSAIS"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   4
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   14737632
      MPTR            =   1
      MICON           =   "frmMenu.frx":14582
      PICN            =   "frmMenu.frx":1459E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Skin_Button.ctr_Button ctr_Button21 
      Height          =   645
      Left            =   120
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   120
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   1138
      BTYPE           =   3
      TX              =   "  FECHAR O SISTEMA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      MICON           =   "frmMenu.frx":148B8
      PICN            =   "frmMenu.frx":148D4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Skin_Button.ctr_Button cmdAtualizar 
      Height          =   645
      Left            =   120
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   7920
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   1138
      BTYPE           =   3
      TX              =   "ATUALIZAR VERSÃO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
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
      MICON           =   "frmMenu.frx":14BEE
      PICN            =   "frmMenu.frx":14C0A
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
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAtualizar_Click()
    Shell (CurDir & "\atualizar.bat")
    End
    'MsgBox (CurDir & "\atualizar.bat")
End Sub

Private Sub cmdCodMwts_Click()
Dim dtIni, dtFim As Date
        
    frm_ESCOLHA_DATA.Show 1
    
    dtIni = CDate(frm_ESCOLHA_DATA.TXT_DT_INICIAL)
    dtFim = CDate(frm_ESCOLHA_DATA.TXT_DT_FINAL)
   
     
    If de.rscmdComiss_Grouping.State = 1 Then de.rscmdComiss_Grouping.Close
    
    de.cmdCod dtIni, dtFim
    rptCodOficial.Show
End Sub

Private Sub cmdGridCaixa_Click()
    frm_GRID_Caixa.Show 1
End Sub

Private Sub cmdGridGerente_Click()
    frm_GRID_Gerente.Show 1
End Sub

Private Sub cmdGridLoja_Click()
    frm_GRID_Loja.Show 1
End Sub

Private Sub cmdGridMeta_Click()
    frm_GRID_Meta.Show 1
End Sub

Private Sub cmdGridVendedor_Click()
    frm_GRID_Vendedor.Show 1
End Sub

Private Sub cmdRelFixosSaldos_Click()
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

Private Sub cmdRelSalarios_Click()
Dim ano, mes As String
      
     ano = InputBox("Entre com o ano", , Format(Date, "YYYY"))
     mes = InputBox("Entre com o mês:", , Format(Date, "MM"))
     
    'If de.rscmdSqlConfComPremioPiso_Grouping = 1 Then de.rscmdSqlConfComPremioPiso_Grouping.Close
     
    de.cmdSqlConfComPremioPiso_Grouping
    rptConfComPremioPiso.Show
    
End Sub

Private Sub cmdRelSalG_Click()
Dim dtIni, dtFim As Date
    
'If de.rscmdSqlSalarioGerentes.State = 1 Then de.rscmdSqlSalarioGerentes.Close
    
    frm_ESCOLHA_DATA.Show 1
    
    dtIni = CDate(frm_ESCOLHA_DATA.TXT_DT_INICIAL)
    dtFim = CDate(frm_ESCOLHA_DATA.TXT_DT_FINAL)
    
     'dtIni = InputBox("Entre com a Data Inicial:", , Format(Date, "DD/MM/YYYY"))
     'dtFim = InputBox("Entre com a Data Final:", , Format(Date, "DD/MM/YYYY"))
     
    If de.rscmdSqlSalarioGerentes.State = 1 Then de.rscmdSqlSalarioGerentes.Close
     
    de.cmdSqlSalarioGerentes dtIni, dtFim
    If IsDate(dtIni) And IsDate(dtFim) Then
        rptSalarioGerentes.Sections("SecCab").Controls("lbTitulo").Caption = "SAL. G (" & Month(dtIni) & ")"
        'rptSalarioGerentes.Sections("SecCab").Controls("lbData").Caption = Format(Date, "DD=MM") & " " & Format(Time, "hh=mm")
         
        rptSalarioGerentes.Show
    End If
End Sub

Private Sub cmdSalarios_Click()

    If MsgBox("Deseja exibir referente a apenas um único mês?", vbYesNo, "Tipo do Relatório") = vbYes Then
        GoTo unico
    Else
        GoTo anual
    End If
        

unico:
     w_mes = InputBox("Entre com o Mês:", "Mês", Format(Date, "MM"))
     w_ano = InputBox("Entre com o Ano:", "Ano", Format(Date, "YYYY"))
     If IsNumeric(w_mes) And IsNumeric(w_ano) Then
        If de.rscmdRelSalarios_Grouping.State = 1 Then de.rscmdRelSalarios_Grouping.Close
         de.cmdRelSalarios_Grouping w_ano, w_mes
         rptSalariosBruto.Show
     Else
        MsgBox "Redigite o Mês e Ano Desejado!", vbExclamation
        GoTo unico
     End If
     GoTo sair
     
anual:
     w_loja = InputBox("Entre com o número da loja:", "Loja")
     w_ano = InputBox("Entre com o Ano:", "Ano", Format(Date, "YYYY"))
     If IsNumeric(w_loja) And IsNumeric(w_ano) Then
        If de.rscmdRelSalariosPorMes_Grouping.State = 1 Then de.rscmdRelSalariosPorMes_Grouping.Close
         de.cmdRelSalariosPorMes_Grouping w_ano, w_loja
         rptSalariosBrutoPorMes.Show
     Else
        MsgBox "Redigite a Loja e o Ano Desejado!", vbExclamation
        GoTo anual
     End If
     GoTo sair
     
     
sair:
    Exit Sub
err1:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume sair

End Sub

Private Sub cmdSemRegistro_Click()

     If de.rsrelSemRegistro.State = 1 Then de.rsrelSemRegistro.Close

ini:
     w_mes = InputBox("Entre com o Mês:", , Format(Date, "MM"))
     w_ano = InputBox("Entre com o Ano:", , Format(Date, "YYYY"))
     If IsNumeric(w_mes) And IsNumeric(w_ano) Then
         de.relSemRegistro w_mes, w_ano
         rptSemRegistro.Show
     Else
        MsgBox "Redigite o Mês e Ano Desejado!", vbExclamation
        GoTo ini
     End If
     
sair:
    Exit Sub
err1:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume sair

End Sub

Private Sub ctr_Button10_Click()
On Error GoTo err1

     If de.rscmdQtde_Func_Logo_Grouping.State = 1 Then de.rscmdQtde_Func_Logo_Grouping.Close
ini:

     w_mes = InputBox("Entre com o Mês:", , Format(Date, "MM"))
     w_ano = InputBox("Entre com o Ano:", , Format(Date, "YYYY"))
     If IsNumeric(w_mes) And IsNumeric(w_ano) Then
     
         de.cmdQtde_Func_Logo_Grouping w_ano, w_mes
         rptRelQtdeEmp.Sections("SecCab").Controls.Item("LbPer").Caption = "Período : " & w_mes & "/" & w_ano
         
         If UCase(InputBox("Mostrar Nomes ?" & Chr(13) & Chr(13) & "S - Sim" & Chr(13) & "N - Não", "Opção", "S")) = "N" Then
              rptRelQtdeEmp.Sections("SecDet").Visible = False
              rptRelQtdeEmp.Sections("SecCabG").Controls.Item("Fundo").BackColor = &HFFFFFF
              rptRelQtdeEmp.Sections("SecCabG").Controls.Item("LB1").Visible = False
              rptRelQtdeEmp.Sections("SecCabG").Controls.Item("LB2").Visible = False
              rptRelQtdeEmp.Sections("SecCabG").Controls.Item("LB3").Visible = False
              rptRelQtdeEmp.Sections("SecCabG").Controls.Item("LB4").Visible = False
              rptRelQtdeEmp.Sections("SecCabG").Controls.Item("lblSal13").Visible = False
              rptRelQtdeEmp.Sections("SecCabG").Controls.Item("lblCalc13").Visible = False
         
         End If
         rptRelQtdeEmp.Show
     
     Else
        MsgBox "Redigite o Mês e Ano Desejado!", vbExclamation
        GoTo ini
     End If


sair:
    Exit Sub
err1:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume sair

End Sub

Private Sub ctr_Button11_Click()
    frm_Escolha_Comp.Show 1
End Sub

Private Sub ctr_Button12_Click()
On Error GoTo err1
 
    FRM_IMP_F.Show 1
     
    w_mes = FRM_IMP_F.TXT_MES
    w_ano = FRM_IMP_F.TXT_ANO
    w_Nome = FRM_IMP_F.dbNome
    w_logo = FRM_IMP_F.TXT_LOGO
    
    If de.rscmdSqlResumoContasLg_Grouping.State = 1 Then de.rscmdSqlResumoContasLg_Grouping.Close
    de.cmdSqlResumoContasLg_Grouping w_mes, w_ano, w_logo
    rptRelResumoContasLgDet.Show
    
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub ctr_Button13_Click()
On Error GoTo err1
 
    FRM_IMP_F.dbNome.Visible = False
    FRM_IMP_F.ck_Nome.Visible = False
    FRM_IMP_F.lbNome.Visible = False
    FRM_IMP_F.Show 1
    
     
    w_mes = FRM_IMP_F.TXT_MES
    w_ano = FRM_IMP_F.TXT_ANO
    w_Nome = FRM_IMP_F.dbNome
    w_logo = FRM_IMP_F.TXT_LOGO
     
     
    If de.rscmdSqlResumoContasLgSINT.State = 1 Then de.rscmdSqlResumoContasLgSINT.Close
    'de.cmdSqlResumoContasLgSINT w_mes, w_Ano, w_logo
    
    w_Sql = "SELECT TAB_FICHA_MENS.M_MES, TAB_FICHA_MENS.M_ANO, TAB_FICHA_MENS.M_LOGO, TAB_LIQ.Total  " & _
            "AS LIQ, TAB_BRUTO.Total as BRUTO, BRUTO + LIQ AS SALDO " & _
            "FROM TAB_FICHA_MENS, TAB_DESC_CALC, " & _
            "(SELECT TAB_FICHA_MENS.M_MES AS mes , TAB_FICHA_MENS.M_ANO AS ano , TAB_FICHA_MENS.M_LOGO AS logo ,  " & _
            "SUM(TAB_DESC_CALC.C_VALOR) AS Total FROM TAB_FICHA_MENS , TAB_DESC_CALC WHERE TAB_FICHA_MENS.M_NFICHA =  " & _
            "TAB_DESC_CALC.C_N_FICHA AND (TAB_DESC_CALC.C_TP_OP <> '=') AND (TAB_FICHA_MENS.M_BLOQ = 0) and (TAB_DESC_CALC.C_VALOR < 0)   " & _
            "GROUP BY TAB_FICHA_MENS.M_MES , TAB_FICHA_MENS.M_ANO , TAB_FICHA_MENS.M_LOGO) TAB_LIQ ,   " & _
            "(SELECT TAB_FICHA_MENS.M_MES AS mes , TAB_FICHA_MENS.M_ANO AS ano , TAB_FICHA_MENS.M_LOGO AS logo ,  " & _
            "SUM(TAB_DESC_CALC.C_VALOR) AS Total FROM TAB_FICHA_MENS , TAB_DESC_CALC WHERE TAB_FICHA_MENS.M_NFICHA =  " & _
            "TAB_DESC_CALC.C_N_FICHA AND (TAB_DESC_CALC.C_TP_OP <> '=') AND (TAB_FICHA_MENS.M_BLOQ = 0) and (TAB_DESC_CALC.C_VALOR > 0)  " & _
            "GROUP BY TAB_FICHA_MENS.M_MES , TAB_FICHA_MENS.M_ANO , TAB_FICHA_MENS.M_LOGO) TAB_BRUTO  " & _
            "Where TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA  " & _
            "And (TAB_FICHA_MENS.M_MES = TAB_LIQ.mes AND TAB_FICHA_MENS.M_MES = TAB_BRUTO.mes) " & _
            "AND (TAB_FICHA_MENS.M_ANO = TAB_LIQ.ano AND TAB_FICHA_MENS.M_ANO = TAB_BRUTO.ano) " & _
            "AND (TAB_FICHA_MENS.M_LOGO = TAB_LIQ.logo AND TAB_FICHA_MENS.M_LOGO = TAB_BRUTO.logo) " & _
            "AND (TAB_DESC_CALC.C_TP_OP <> '=')  " & _
            "GROUP BY TAB_FICHA_MENS.M_MES, TAB_FICHA_MENS.M_ANO, TAB_FICHA_MENS.M_LOGO, TAB_LIQ.Total,TAB_BRUTO.Total " & _
            "HAVING (TAB_FICHA_MENS.M_MES = " & w_mes & ") AND (TAB_FICHA_MENS.M_ANO = " & w_ano & ") AND  " & _
            "(TAB_FICHA_MENS.M_LOGO LIKE '" & w_logo & "') ORDER BY TAB_FICHA_MENS.M_MES, TAB_FICHA_MENS.M_ANO"

    de.rscmdSqlResumoContasLgSINT.Open w_Sql
    
    rptRelResumoContasLg.Sections(2).Controls("lbPer").Caption = "  Período :  " & Format(w_mes, "00") & " / " & w_ano
    rptRelResumoContasLg.Show
    
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub ctr_Button14_Click()
On Error GoTo err1

 FRM_IMP_F.dbNome.Visible = False
 FRM_IMP_F.lbNome.Visible = False
 FRM_IMP_F.ck_Nome.Visible = False
 
 FRM_IMP_F.Show 1
 
w_mes = FRM_IMP_F.TXT_MES
w_ano = FRM_IMP_F.TXT_ANO
w_logo = FRM_IMP_F.TXT_LOGO & "%"
    
If FRM_IMP_F.txt_State = "A" And IsNumeric(w_mes) And IsNumeric(w_ano) Then
    
    If de.rscmdSqlTP.State = 1 Then de.rscmdSqlTP.Close
    de.cmdSqlTP w_mes, w_ano, w_logo
    
    If Not de.rscmdSqlTP.EOF Then
        
        rptRelTP.Sections("seccab").Controls("lbPer").Caption = "  Período :  " & Format(w_mes, "00") & " / " & w_ano
        
        wTot = 0
        
        Do While Not de.rscmdSqlTP.EOF
            wTot = wTot + CDbl(de.rscmdSqlTP.Fields("TOTAL_TP"))
            de.rscmdSqlTP.MoveNext
        Loop
        
        rptRelTP.Sections("secrod").Controls("lbTot").Caption = Format(wTot, "0")
        rptRelTP.Show
        
    Else
        MsgBox "NÃO EXISTE T.P NESTE PERÍODO : " & w_mes & "/" & w_ano, vbInformation
    End If
        
    W_CONT = 0
Else
    MsgBox "Relatório Cancelado!", vbInformation
End If
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub ctr_Button15_Click()
    rptEmprestimos.Show
End Sub

Private Sub ctr_Button16_Click()
    rptEmprestimoAnalise.Show
End Sub

Private Sub ctr_Button17_Click()
    FRM_IMP_F.Show 1
     
    'w_mes = FRM_IMP_F.TXT_MES
    'w_ano = FRM_IMP_F.TXT_ANO
    'w_Nome = FRM_IMP_F.dbNome
    'w_logo = FRM_IMP_F.TXT_LOGO
      
    de.sqlComissaoPremio FRM_IMP_F.TXT_MES, FRM_IMP_F.TXT_ANO, FRM_IMP_F.TXT_LOGO
    rptComissaoPremio.Show
    
        
    
    
    
    'FRM_IMP_F.Show 1
     
    'w_mes = FRM_IMP_F.TXT_MES
    'w_ano = FRM_IMP_F.TXT_ANO
    'w_Nome = FRM_IMP_F.dbNome
    'w_logo = FRM_IMP_F.TXT_LOGO
    
    
    'wSQL = " SHAPE {SELECT * FROM `Con_Rpt_Com_Vendas` " & _
           " WHERE (M_LOGO LIKE '" & w_logo & "') AND (M_MES = " & w_mes & ") AND (M_ANO = " & w_ano & ") AND (F_NOME LIKE '" & w_Nome & "')" & _
           "}  AS Con_Rpt_Com_Vendas COMPUTE Con_Rpt_Com_Vendas BY 'M_LOGO'"
    
    'If de.rsCon_Rpt_Com_Vendas_Grouping.State = 1 Then de.rsCon_Rpt_Com_Vendas_Grouping.Close
    'de.rsCon_Rpt_Com_Vendas_Grouping.Open wSQL

    'rptVendasCom.Sections("Cab").Controls("lbTitulo").Caption = " Ref.  " & w_mes & "/" & w_ano
    'rptVendasCom.Show
End Sub

Private Sub ctr_Button18_Click()
On Error GoTo err1

     If de.rscmdSqlSalarioCX_Grouping.State = 1 Then de.rscmdSqlSalarioCX_Grouping.Close
ini:

     w_mes = InputBox("Entre com o Mês:", , Format(Date, "MM"))
     w_ano = InputBox("Entre com o Ano:", , Format(Date, "YYYY"))
     If IsNumeric(w_mes) And IsNumeric(w_ano) Then
     
         de.cmdSqlSalarioCX_Grouping w_mes, w_ano
         
         rptSalarioCx.Show
     
     Else
        MsgBox "Redigite o Mês e Ano Desejado!", vbExclamation
        GoTo ini
     End If


sair:
    Exit Sub
err1:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub ctr_Button19_Click()
    frm_Log.Show 1
End Sub

Private Sub ctr_Button20_Click()
    frm_Emprest.Show 1
End Sub

Private Sub ctr_Button21_Click()
    If vbYes = MsgBox("Deseja realmente Sair?", vbQuestion + vbYesNo) Then End
End Sub

Private Sub ctr_Button22_Click()

End Sub

Private Sub ctr_Button9_Click()
    Backup
End Sub

Private Sub btnFichasMensais_Click()
    frm_Alt_Fic_Mensal_VIS.Show
End Sub

Private Sub cmdAddLanç_SalF_Click()
    frm_Cad_Funcionario.Show 1
End Sub

Private Sub ctr_Button1_Click()
    frm_Alt_Funcionario.Show 1
End Sub

Private Sub ctr_Button2_Click()
    frm_Cad_Tp_Conta.Show 1
End Sub

Private Sub ctr_Button3_Click()
    frm_Alt_TP_CONTA.Show 1
End Sub

Private Sub ctr_Button4_Click()
    frm_Vendas.Show 1
End Sub

Private Sub ctr_Button5_Click()
    frm_Cad_Sal_Familia.Show 1
End Sub

Private Sub ctr_Button6_Click()
    frm_Acesso_Especial.Show 1
End Sub

Private Sub ctr_Button7_Click()
    frm_Gerar_Fichas.Show 1
End Sub

Private Sub ctr_Button8_Click()
    frm_Alt_Visto_Vale.Show 1
End Sub

Private Sub Form_Activate()
    If acessoTotal = False Then
        cmdGridGerente.Enabled = False
        cmdGridLoja.Enabled = False
        cmdGridCaixa.Enabled = False
        cmdRelSalG.Enabled = False
        cmdSalarios.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    If w_umaVez = 0 Then
        frmSplash.PB.value = 20
        frmSplash.PB.text = "Carregando " & frmSplash.PB.value & "%"
    End If
    frm_Alt_Fic_Mensal_VIS.Show
End Sub

