VERSION 5.00
Begin VB.Form frm_Habilitar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Senha Mestre"
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   2460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_Pss 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   1800
      Picture         =   "frm_Habilitar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   495
      Left            =   1320
      Picture         =   "frm_Habilitar.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frm_Habilitar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    txt_Pss = ""
    Hide
    
End Sub

Private Sub cmdOK_Click()
    txt_Pss.Text = md5Test.DigestStrToHexStr(LCase(txt_Pss.Text))
    Hide
End Sub

Private Sub Form_Activate()
    txt_Pss = ""
    txt_Pss.SetFocus
End Sub

