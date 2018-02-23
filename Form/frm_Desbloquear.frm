VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_Desbloquear 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DESBLOQUEAR FICHAS"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6240
   Icon            =   "frm_Desbloquear.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ckTodos 
      Caption         =   "Selecionar Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   480
      Width           =   1815
   End
   Begin VB.ListBox TXT_LOGO 
      Height          =   2400
      Left            =   600
      MultiSelect     =   1  'Simple
      TabIndex        =   16
      Top             =   360
      Width           =   735
   End
   Begin MSDataListLib.DataList TXT_LOGO3 
      Bindings        =   "frm_Desbloquear.frx":12D2
      DataSource      =   "ADO_CENTRAL"
      Height          =   450
      Left            =   2880
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   794
      _Version        =   393216
      MatchEntry      =   -1  'True
      ListField       =   "LOJA"
      BoundColumn     =   ""
      Object.DataMember      =   "TAB_L"
   End
   Begin VB.CheckBox ck_Nome 
      Caption         =   "Todos Nomes"
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
      Left            =   4440
      TabIndex        =   5
      Top             =   1905
      Width           =   1560
   End
   Begin VB.CheckBox ckTodas 
      Caption         =   "Todos Logos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt_State 
      Height          =   285
      Left            =   3360
      TabIndex        =   13
      Text            =   "F"
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdImp 
      Caption         =   "&Desbloquear"
      Height          =   735
      Left            =   4440
      Picture         =   "frm_Desbloquear.frx":12E4
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCanc 
      Caption         =   "&Cancelar"
      Height          =   735
      Left            =   1080
      Picture         =   "frm_Desbloquear.frx":1726
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3120
      Width           =   735
   End
   Begin MSDataListLib.DataCombo dbNome 
      Bindings        =   "frm_Desbloquear.frx":1A30
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   2160
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      ListField       =   "F_NOME"
      Text            =   "%"
      Object.DataMember      =   "TAB_FUNCIONARIO"
   End
   Begin VB.ComboBox TXT_MES 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frm_Desbloquear.frx":1A41
      Left            =   3960
      List            =   "frm_Desbloquear.frx":1A69
      TabIndex        =   2
      Top             =   1200
      Width           =   780
   End
   Begin VB.TextBox TXT_ANO 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4965
      TabIndex        =   3
      Top             =   1200
      Width           =   810
   End
   Begin MSDataListLib.DataCombo TXT_LOGO2 
      Bindings        =   "frm_Desbloquear.frx":1A94
      DataField       =   "F_COD_L"
      DataSource      =   "ADOREG"
      Height          =   360
      Left            =   1920
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      ListField       =   "COD_LOJ"
      BoundColumn     =   "COD_LOJ"
      Text            =   "%"
      Object.DataMember      =   "TAB_L"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADO_CENTRAL 
      Height          =   330
      Left            =   2160
      Top             =   3600
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
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
      BackColor       =   -2147483643
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
      Caption         =   "CENTRAL"
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Para todos coloque  ""%"""
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   " Opções : "
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
      Left            =   255
      TabIndex        =   12
      Top             =   45
      Width           =   975
   End
   Begin VB.Shape Shape2 
      Height          =   975
      Left            =   120
      Top             =   3000
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      Height          =   2775
      Left            =   120
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label lbNome 
      BackStyle       =   0  'Transparent
      Caption         =   "NOME"
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
      Left            =   1440
      TabIndex        =   11
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "MÊS"
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
      Left            =   4080
      TabIndex        =   10
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ANO"
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
      Left            =   5160
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "(B)"
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
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frm_Desbloquear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset

Private Sub ck_Nome_Click()
    If ck_Nome.value = 1 Then
        dbNome = "%"
        dbNome.Enabled = False
    Else
        dbNome = ""
        dbNome.Enabled = True
        dbNome.SetFocus
        SendKeys "{f4}"
    End If
End Sub

Private Sub ck_Nome_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub ckTodas_Click()
    If ckTodas.value = 1 Then
        TXT_LOGO = "%"
        'TXT_LOGO.Enabled = False
    Else
        TXT_LOGO = ""
        'TXT_LOGO.Enabled = True
        TXT_LOGO.SetFocus
        SendKeys "{f4}"
    End If
End Sub

Private Sub ckTodas_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub ckTodos_Click()
    If ckTodos.value = 1 Then
        ckTodas.value = 1
        For I = 0 To TXT_LOGO.ListCount - 1
            TXT_LOGO.Selected(I) = True
        Next
    Else
        ckTodas.value = 0
        For I = 0 To TXT_LOGO.ListCount - 1
            TXT_LOGO.Selected(I) = False
        Next
    End If
    Call ckTodas_Click
End Sub

Private Sub cmdCanc_Click()
On Error Resume Next
    txt_State = "F"
    Unload Me
End Sub

Private Sub cmdImp_Click()
Dim nenhumaLoja As Boolean

On Error Resume Next

    nenhumaLoja = True
    For I = 0 To TXT_LOGO.ListCount - 1
        If TXT_LOGO.Selected(I) = True Then
           nenhumaLoja = False
        End If
    Next

    If nenhumaLoja Then
        MsgBox "Selecione ao menos um (B)!", vbCritical
    Else
        txt_State = "A"
        FRM_IMP_F.dbNome.Visible = True
        FRM_IMP_F.ck_Nome.Visible = True
        FRM_IMP_F.lbNome.Visible = True
        Hide
    End If
End Sub

Private Sub dbNome_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub Form_Activate()
    For I = 0 To TXT_LOGO.ListCount - 1
        If TXT_LOGO.List(I) = frm_Alt_Fic_Mensal_VIS.txtLogo.Text Then
           TXT_LOGO.Selected(I) = True
        End If
    Next
    
    
    
    
End Sub

Private Sub Form_Load()
On Error Resume Next
    
    TXT_LOGO = "%"
    TXT_ANO = Format(Date, "yyyy")
    TXT_MES = Format(Date, "mm")
    dbNome = "%"
    Set ADO_CENTRAL.Recordset = de.cnc.Execute("SELECT COD_LOJ FROM LOJB010 ORDER BY COD_LOJ").Clone
    Set rs = ADO_CENTRAL.Recordset.Clone
    rs.MoveFirst
    Do While Not rs.EOF
        TXT_LOGO.AddItem (rs("COD_LOJ"))
        rs.MoveNext
    Loop
    

End Sub



'KeyPress
Private Sub dbNome_KeyPress(KeyAscii As Integer)
    KeyEnter KeyAscii
End Sub

Private Sub TXT_ANO_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub TXT_ANO_KeyPress(KeyAscii As Integer)
    KeyEnter KeyAscii
End Sub

Private Sub TXT_LOGO_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub TXT_LOGO_KeyPress(KeyAscii As Integer)
    KeyEnter KeyAscii
End Sub

Private Sub TXT_MES_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub TXT_MES_KeyPress(KeyAscii As Integer)
    KeyEnter KeyAscii
End Sub
