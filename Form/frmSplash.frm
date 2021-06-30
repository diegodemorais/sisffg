VERSION 5.00
Object = "{9A4D18F7-4EC7-11D5-9E33-0040C78773FC}#1.0#0"; "VBxPOLITEC.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4470
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":548A
   ScaleHeight     =   4470
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timInicio 
      Interval        =   1
      Left            =   6840
      Top             =   120
   End
   Begin VBXPolitec.ocxProgressBarTexto PB 
      Height          =   360
      Left            =   240
      TabIndex        =   5
      Top             =   3960
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorFundo  =   16711680
      BackColorFundo  =   -2147483643
      BackColorProgress=   16711680
      MaxProgress     =   100
   End
   Begin VB.Label lblComments 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Comentários"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   660
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   4065
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fichas de Funcionários"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   270
      TabIndex        =   6
      Top             =   220
      Width           =   6075
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Desenvolvido para:"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "RP Assessoria"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Versão : 1.0"
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
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   2760
      Width           =   1350
   End
   Begin VB.Label lblWarning 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   3960
      Width           =   4815
   End
   Begin VB.Label lblProductName 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fichas de Funcionários"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6075
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w_time As Byte

Option Explicit

Private Sub Form_Load()
    lblVersion.Caption = "Versão: " & App.Major & "." & App.Minor & "." & App.Revision
    lblComments.Caption = App.Comments
    'lblProductName.Caption = App.Title
    'frmSplash.lblWarning = frmSplash.lblWarning & " **********   BASE OFICIAL !  ***********"
       
        
    timInicio.Enabled = True 'ligar atualização
    
End Sub



Sub timInicio_Timer()
On Error GoTo err1
    
    'instanciação / atualização
    timInicio.Enabled = False
    
    'AtualizarGeral 'abrir as tabelas e instanciar objetos
    'frm_Alt_Fic_Mensal_VIS.Show
    
sair:

    PB.value = 10
    PB.text = "Carregando " & frmSplash.PB.value & "%"
    
    mdiPrincipal.Show
    frmSplash.Hide
    Unload Me
    
    Exit Sub
err1:

If CDbl(Err.Number) = CDbl(424) Then
        'If CDbl(Err.Number) = CDbl(424) And Not (UCase(frmLogin.txtUserName.Text) = NomeMestre) Then
        'MsgBox "Favor fazer a importação de tabelas!", vbCritical
    Else
         MsgBox Error$, vbCritical
    End If
    Resume sair
    
End Sub

