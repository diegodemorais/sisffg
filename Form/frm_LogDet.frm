VERSION 5.00
Begin VB.Form frm_LogDet 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DETALHE do Registro"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   9120
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   7560
      TabIndex        =   16
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtDescricaoOLD 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3240
      Width           =   8895
   End
   Begin VB.TextBox txtDescricaoNEW 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1680
      Width           =   8895
   End
   Begin VB.TextBox txtTabela 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox txtAcao 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtUsuario 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtHora 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtData 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtCod 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblDescricaoOLD 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição ANTIGA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição NOVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Tabela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   10
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Hora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "frm_LogDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

