VERSION 5.00
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fichas de Funcionários [SisFF]"
   ClientHeight    =   1995
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4080
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1178.713
   ScaleMode       =   0  'User
   ScaleWidth      =   3830.899
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   0
      Text            =   "RP"
      Top             =   240
      Width           =   1485
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   1485
   End
   Begin Skin_Button.ctr_Button cmdOK 
      Height          =   525
      Left            =   360
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   "&ENTRAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "frmLogin.frx":0442
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Skin_Button.ctr_Button cmdCancel 
      Height          =   525
      Left            =   2160
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   "&SAIR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "frmLogin.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   720
      TabIndex        =   4
      Top             =   360
      Width           =   840
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   720
      TabIndex        =   5
      Top             =   720
      Width           =   840
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'= Global Variables



Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter (KeyCode)
    If KeyCode = vbKeyEscape Then cmdCancel_Click
End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter (KeyCode)
    If KeyCode = vbKeyEscape Then cmdCancel_Click
End Sub

Private Sub Form_Activate()
    txtPassword.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter (KeyCode)
    If KeyCode = vbKeyEscape Then cmdCancel_Click
End Sub

Private Sub Form_Load()
    Dim strUser

    ' Instantiate our class
    Set md5Test = New MD5
    strUser = GetIni("SYSTEM", "User", App.Path & "\System.INI")
    If strUser <> "" Then
        txtUserName = strUser
    End If
    
        
End Sub

'*** Botões ***

Sub cmdOK_Click()
    'Pega os Dir. do Arq. INI
    strDirBase = GetIni("SYSTEM", "DirBase", App.Path & "\System.INI")
    strDirBaseCentral = GetIni("SYSTEM", "DirBaseCentral", App.Path & "\System.INI")
    strDirBaseServer = GetIni("SYSTEM", "DirBaseServer", App.Path & "\System.INI")
    strDirRPT = GetIni("SYSTEM", "DirRPT", App.Path & "\System.INI")
    strImpressora = GetIni("SYSTEM", "Impressora", App.Path & "\System.INI")
    strImgFundo = GetIni("SYSTEM", "ImgFundo", App.Path & "\System.INI")
    strImgSplash = GetIni("SYSTEM", "ImgSplash", App.Path & "\System.INI")
    
    
     'Passa a string de conecção com o Diretorio do Arq. INI
     If de.cnc.State = 0 Then de.cnc.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & strDirBase & ";Mode=Share Deny None;Persist Security Info=False"
     'If de.cnc.State = 0 Then de.cnc.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & strDirBase & ";Mode=Share Deny None;Persist Security Info=False"
     
     'If de.cnc.State = 0 Then de.cnc.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & strDirBase & ";Mode=Share Deny None;Persist Security Info=False;Jet OLEDB:Database Password=poter12"
     'If de.cncPDX.State = 0 Then de.cncPDX.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDirBaseCentral & ";Extended Properties=Paradox 5.x;Persist Security Info=False"
     If de.cncDBase.State = 0 Then de.cncDBase.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & strDirBase & ";Mode=Share Deny None;Persist Security Info=False"   '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDirBaseCentral & ";Extended Properties=DBase IV;Persist Security Info=False"
     'If de.cncDBase.State = 0 Then de.cncDBase.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & strDirBase & ";Mode=Share Deny None;Persist Security Info=False;Jet OLEDB:Database Password=poter12"   '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDirBaseCentral & ";Extended Properties=DBase IV;Persist Security Info=False"
                                                                  
     If de.rsTab_Config.State = 0 Then de.Tab_Config
     If Not de.rsTab_Config.EOF Then
        w_PassWordLib = de.rsTab_Config.Fields("pss_liberacao")
        SenhaMestre = de.rsTab_Config.Fields("pss_login")
     Else
        w_PassWordLib = ""
        SenhaMestre = ""
     End If
     
     
                                                                  
    p_Usuario = txtUserName
    w_usuario = txtUserName
    w_usuario2 = w_usuario
                                                                  
   'Check de Senha e Nome é do Comum
    If (txtPassword = SenhaUsu And UCase(txtUserName) = NomeUsu) Or (txtPassword = SenhaUsu2 And UCase(txtUserName) = NomeUsu2) Or (txtPassword = SenhaUsu3 And UCase(txtUserName) = NomeUsu3) Or (txtPassword = SenhaUsu4 And UCase(txtUserName) = NomeUsu4) Or (txtPassword = SenhaUsu5 And UCase(txtUserName) = NomeUsu5) Then
        w_usuario = txtUserName
         On Error Resume Next
        frmSplash.Picture = LoadPicture(strImgSplash)

      
      'Desabilita os menus
           
       'Fecha o Form
         Hide 'Unload Me
       'Abre a tela de splash
         frmSplash.Show


    'Mudar para False
        'mdiPrincipal.Picture = LoadPicture(strImgFundo)
        'mdiPrincipal.mnuMaster.Enabled = True
        'mdiPrincipal.mnuSisMenVisVal.Enabled = False
        'mdiPrincipal.mnuLog.Visible = False
         

     
   'Check senha e Nome é do Mestre
    ElseIf md5Test.DigestStrToHexStr(LCase(txtPassword.text)) = UCase(SenhaMestre) And (UCase(txtUserName.text) = UCase(NomeMestre) Or UCase(txtUserName.text) = UCase(NomeMestre2) Or UCase(txtUserName.text) = UCase(NomeMestre3)) Then
        
        'Habilita os menus
        frmSplash.Picture = LoadPicture(strImgSplash)
       'Abre a tela de splash
         frmSplash.Show
         
       'Fecha o Form
         Hide  'Unload Me

         
         
        'mdiPrincipal.Picture = LoadPicture(strImgFundo)
        'mdiPrincipal.mnuMaster.Enabled = True
        'mdiPrincipal.mnuLog.Visible = True
        
        
   'Se não for nenhuma das comparações
    Else
        MsgBox "Usuário ou Senha Inválidos, tente novamente!", , "Sistema de Fichas [SisFF] - Erro de autenticação"
        txtPassword.SetFocus
        Sendkeys "{Home}+{End}"
    End If

End Sub

Private Sub cmdCancel_Click()
    End
End Sub




Private Sub txtPassword_GotFocus()
    Sendkeys "{home}+{end}"
End Sub

'*** Keydowns das txt
Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter (KeyCode)
    If KeyCode = 13 Then cmdOK_Click
    If KeyCode = vbKeyEscape Then cmdCancel_Click
End Sub

Private Sub txtUserName_GotFocus()
    Sendkeys "{home}+{end}"
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter (KeyCode)
    If KeyCode = vbKeyEscape Then cmdCancel_Click
End Sub
