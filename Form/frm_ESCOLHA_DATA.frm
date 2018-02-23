VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_ESCOLHA_DATA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ESCOLHA ESCOPO DE DATA"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "frm_ESCOLHA_DATA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker TXT_DT_INICIAL 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   71106561
      CurrentDate     =   38282
   End
   Begin MSComCtl2.DTPicker TXT_DT_FINAL 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   71106561
      CurrentDate     =   38282
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ESCOLHA_DATA.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ESCOLHA_DATA.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ESCOLHA_DATA.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ESCOLHA_DATA.frx":0C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ESCOLHA_DATA.frx":0F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ESCOLHA_DATA.frx":128C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Skin_Button.ctr_Button cmdOk 
      Height          =   645
      Left            =   3840
      TabIndex        =   2
      Top             =   240
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   1138
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
      MICON           =   "frm_ESCOLHA_DATA.frx":15A6
      PICN            =   "frm_ESCOLHA_DATA.frx":15C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA FINAL"
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
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA INICIAL"
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
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label7 
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
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   120
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frm_ESCOLHA_DATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Fechar()
On Error GoTo err1
    Hide
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub cmdAddLanç_SalF_Click()

End Sub

Private Sub cmdOK_Click()
    Fechar
End Sub

Private Sub Form_Activate()
On Error GoTo err1
    TXT_DT_INICIAL = Date
    TXT_DT_FINAL = Date
        
    TXT_DT_INICIAL.SetFocus
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub TXT_DT_FINAL_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub TXT_DT_INICIAL_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub

'------ Ao Pressionar Tecla --------

Private Sub TXT_DT_INICIAL_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub TXT_DT_Final_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

' -------  Teclas de Atalhos --------
Sub Keys(KeyCode As Integer, Shift As Integer)
    '*** Shift (4 = Alt) ***
    If Shift = 4 Then
        Select Case KeyCode
        Case 70: ' "F"
                Fechar
        End Select
    End If
End Sub
