VERSION 5.00
Begin VB.Form frm_GRID_Caixa_Det 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DETALHE do Registro"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   9150
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNotasOld 
      Height          =   285
      Left            =   3360
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txtFicha 
      Height          =   285
      Left            =   6360
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFunc 
      Height          =   285
      Left            =   5040
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtNotas 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   8895
   End
   Begin VB.TextBox txtNome 
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   360
      Width           =   4695
   End
   Begin VB.TextBox txtTipo 
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtB 
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtSigla 
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Notas"
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
      TabIndex        =   10
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
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
      Left            =   2880
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cx"
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
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "B"
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
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sigla"
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
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frm_GRID_Caixa_Det"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub cmdCancelar_Click()
    cmdSalvar_Click
End Sub

Sub cmdSalvar_Click()
    If txtNotas <> txtNotasOld Then
        If vbYes = MsgBox("Deseja SALVAR as Notas?", vbQuestion + vbYesNo) Then
            de.cnc.Execute ("UPDATE TAB_FUNCIONARIO SET F_NOTAS = '" & txtNotas & "' WHERE F_CODIGO = " & txtFunc)
            de.cnc.Execute ("UPDATE TAB_FICHA_MENS SET M_NOTAS = '" & txtNotas & "' WHERE M_NFICHA = " & txtFicha)
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Deactivate()
    cmdCancelar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then cmdCancelar_Click
   If KeyAscii = vbEnter Then cmdSalvar_Click
End Sub

Private Sub Form_LostFocus()
    cmdCancelar_Click
End Sub

Private Sub txtNotas_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then cmdCancelar_Click
   If KeyAscii = vbEnter Then cmdSalvar_Click
End Sub
