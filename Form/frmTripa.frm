VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frmTripa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualização da Tripinha"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   Icon            =   "frmTripa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   5775
   StartUpPosition =   1  'CenterOwner
   Begin Skin_Button.ctr_Button cmdFechar 
      Height          =   615
      Left            =   4050
      TabIndex        =   2
      Top             =   7200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   "&Fechar"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTripa.frx":12D2
      PICN            =   "frmTripa.frx":12EE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RichTextLib.RichTextBox DocumentoTxt 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   11245
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      FileName        =   "D:\Sistemas\SisFF\Rpts\Tripa.txt"
      TextRTF         =   $"frmTripa.frx":1608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Skin_Button.ctr_Button cmdPrint 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   7200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   "&Imprimir"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTripa.frx":3CFD
      PICN            =   "frmTripa.frx":3D19
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
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Visualização da Tripinha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   1
      Top             =   165
      Width           =   5565
   End
End
Attribute VB_Name = "frmTripa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim a, fs As Object
On Error GoTo err1

    'Cria Arquivo texto para Impressão da Tripa
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.createTextFile(strDirRPT & "\print.bat")
    
    a.Writeline ("@Echo OFF")
    a.Writeline ("Type " & strDirRPT & "\tripa.txt > " & strImpressora)
    a.Writeline ("Exit")
    a.Close
    
    Call Shell(strDirRPT & "\print.bat")

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Form_Load()
On Error GoTo err1
    DocumentoTxt.LoadFile strDirRPT & "\tripa.txt"
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub
