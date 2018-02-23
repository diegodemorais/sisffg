VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRM_IMP_F 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " IMPRESSÃO DE FICHA"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6210
   Icon            =   "FRM_IMP_F.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ckTipo 
      Caption         =   "Todos"
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
      Left            =   2880
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   800
      Width           =   975
   End
   Begin VB.ListBox txt_tipo 
      Enabled         =   0   'False
      Height          =   2010
      ItemData        =   "FRM_IMP_F.frx":12D2
      Left            =   1560
      List            =   "FRM_IMP_F.frx":12F1
      MultiSelect     =   1  'Simple
      TabIndex        =   20
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CheckBox ckTodos 
      Caption         =   "Selecionar todos os (B)"
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
      Left            =   240
      TabIndex        =   18
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox TXT_LOGO 
      Height          =   2790
      Left            =   600
      MultiSelect     =   1  'Simple
      TabIndex        =   17
      Top             =   840
      Width           =   735
   End
   Begin MSDataListLib.DataList TXT_LOGO3 
      Bindings        =   "FRM_IMP_F.frx":1348
      DataSource      =   "ADO_CENTRAL"
      Height          =   450
      Left            =   5160
      TabIndex        =   16
      Top             =   1920
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
   Begin VB.CheckBox CkFicha 
      Caption         =   "Imprimir Ficha?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   4320
      TabIndex        =   15
      Top             =   240
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.CheckBox CkTripa 
      Caption         =   "Imprimir Tripa?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   4320
      TabIndex        =   14
      Top             =   480
      Visible         =   0   'False
      Width           =   1605
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
      Left            =   4320
      TabIndex        =   5
      Top             =   3105
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
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt_State 
      Height          =   285
      Left            =   5760
      TabIndex        =   13
      Text            =   "F"
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdImp 
      Caption         =   "&Imprimir"
      Height          =   735
      Left            =   4440
      Picture         =   "FRM_IMP_F.frx":135A
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdCanc 
      Caption         =   "&Cancelar"
      Height          =   735
      Left            =   1080
      Picture         =   "FRM_IMP_F.frx":179C
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4080
      Width           =   735
   End
   Begin MSDataListLib.DataCombo dbNome 
      Bindings        =   "FRM_IMP_F.frx":1AA6
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   3480
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
      ItemData        =   "FRM_IMP_F.frx":1AB7
      Left            =   3960
      List            =   "FRM_IMP_F.frx":1ADF
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
      Bindings        =   "FRM_IMP_F.frx":1B0A
      DataField       =   "F_COD_L"
      DataSource      =   "ADOREG"
      Height          =   360
      Left            =   4200
      TabIndex        =   0
      Top             =   1920
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
      Top             =   4560
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO"
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
      Left            =   1560
      TabIndex        =   19
      Top             =   840
      Width           =   495
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
      Top             =   3960
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      Height          =   3735
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
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   3120
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
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "FRM_IMP_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset

Private Sub ck_Nome_Click()
    If ck_Nome.value = 1 Then
        dbNome = "%"
        dbNome.Enabled = False
        txt_tipo.Enabled = True
    Else
        dbNome = ""
        dbNome.Enabled = True
        txt_tipo.Enabled = False
        On Error Resume Next
        dbNome.SetFocus
        SendKeys "{f4}"
    End If
End Sub

Private Sub ck_Nome_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub ckTipo_Click()
    If ckTipo.value = 1 Then
        For I = 0 To txt_tipo.ListCount - 1
            txt_tipo.Selected(I) = True
        Next
    Else
        For I = 0 To txt_tipo.ListCount - 1
            txt_tipo.Selected(I) = False
        Next
    End If
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
        ck_Nome.value = 1
        ck_Nome.Enabled = False
        For I = 0 To TXT_LOGO.ListCount - 1
            TXT_LOGO.Selected(I) = True
        Next
    Else
        ckTodas.value = 0
        ck_Nome.value = 0
        ck_Nome.Enabled = True
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
Dim nenhumTipo As Boolean

On Error Resume Next

    nenhumaLoja = True
    For I = 0 To TXT_LOGO.ListCount - 1
        If TXT_LOGO.Selected(I) = True Then
           nenhumaLoja = False
        End If
    Next

    nenhumTipo = True
    For J = 0 To txt_tipo.ListCount - 1
        If txt_tipo.Selected(J) = True Then
           nenhumTipo = False
        End If
    Next
    
     
    If nenhumaLoja Then
        MsgBox "Selecione ao menos um (B)!", vbCritical
        ElseIf nenhumTipo Then
            MsgBox "Selecione ao menos um TIPO!", vbCritical
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

Sub Form_Activate()
    'Limpando
    For I = 0 To TXT_LOGO.ListCount - 1
           TXT_LOGO.Selected(I) = False
    Next
    'Selecionando o LOGO da ficha atual
    For I = 0 To TXT_LOGO.ListCount - 1
        If TXT_LOGO.List(I) = frm_Alt_Fic_Mensal_VIS.txtLogo.Text Then
           TXT_LOGO.Selected(I) = True
        End If
    Next
    
    'Limpando
    For J = 0 To txt_tipo.ListCount - 1
           txt_tipo.Selected(J) = False
    Next
    'Selecionando o TIPO da ficha atual
    For J = 0 To txt_tipo.ListCount - 1
        If txt_tipo.List(J) = frm_Alt_Fic_Mensal_VIS.TXT_FTIPO.Caption Then
           txt_tipo.Selected(J) = True
        End If
    Next
    
    ckTodos.value = False
    ck_Nome.value = False
    
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

