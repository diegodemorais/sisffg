VERSION 5.00
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_Opt_GerenteCaixa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GRID - Gerente ou Caixa"
   ClientHeight    =   2910
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   " OUTROS "
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
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   3975
      Begin Skin_Button.ctr_Button cmdGridMeta 
         Height          =   525
         Left            =   2040
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
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
         MICON           =   "frm_Opt_GerenteCaixa.frx":0000
         PICN            =   "frm_Opt_GerenteCaixa.frx":001C
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
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
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
         MICON           =   "frm_Opt_GerenteCaixa.frx":27CE
         PICN            =   "frm_Opt_GerenteCaixa.frx":27EA
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
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0FF&
      Caption         =   " CARGOS "
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin Skin_Button.ctr_Button cmdGridGerente 
         Height          =   525
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1080
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
         MICON           =   "frm_Opt_GerenteCaixa.frx":4F9C
         PICN            =   "frm_Opt_GerenteCaixa.frx":4FB8
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
         Left            =   2040
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
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
         MICON           =   "frm_Opt_GerenteCaixa.frx":776A
         PICN            =   "frm_Opt_GerenteCaixa.frx":7786
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
         Left            =   120
         TabIndex        =   3
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
         MICON           =   "frm_Opt_GerenteCaixa.frx":9F38
         PICN            =   "frm_Opt_GerenteCaixa.frx":9F54
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
Attribute VB_Name = "frm_Opt_GerenteCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCaixa_Click()
    Hide
    frm_GRID_Caixa.Show
End Sub

Private Sub cmdGerente_Click()
    Hide
    frm_GRID_Gerente.Show
End Sub

Private Sub cmdGridCaixa_Click()
    Hide
    frm_GRID_Caixa.Show
End Sub

Private Sub cmdGridGerente_Click()
    Hide
    frm_GRID_Gerente.Show
End Sub

Private Sub cmdGridLoja_Click()
    Hide
    frm_GRID_Loja.Show
End Sub

Private Sub cmdGridMeta_Click()
    Hide
    frm_GRID_Meta.Show
End Sub

Private Sub cmdGridVendedor_Click()
    Hide
    frm_GRID_Vendedor.Show
End Sub
