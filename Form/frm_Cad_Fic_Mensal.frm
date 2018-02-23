VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_Cad_Fic_Mensal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CADASTRO DE FICHA MENSAL"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "frm_Cad_Fic_Mensal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   5490
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc adoFuncCPF 
      Height          =   375
      Left            =   3240
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Caption         =   "Adodc1"
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
   Begin MSAdodcLib.Adodc adoFuncCod 
      Height          =   375
      Left            =   840
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Caption         =   "Adodc1"
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
   Begin MSAdodcLib.Adodc adoFunc 
      Height          =   375
      Left            =   3240
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Caption         =   "Adodc1"
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
   Begin VB.TextBox TXT_DT_REG 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      DataSource      =   "ADOREG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1740
      MaxLength       =   10
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txt_DT_ADM 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      DataSource      =   "ADOREG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   360
      MaxLength       =   10
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1935
      Width           =   1215
   End
   Begin VB.TextBox txt_ANOTACAO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   6600
      Width           =   4695
   End
   Begin VB.TextBox TXT_ANO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3765
      TabIndex        =   3
      Top             =   1200
      Width           =   810
   End
   Begin MSDataListLib.DataCombo TXT_FUNC 
      Bindings        =   "frm_Cad_Fic_Mensal.frx":1CFA
      Height          =   360
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Locked          =   -1  'True
      MatchEntry      =   -1  'True
      ListField       =   "F_NOME"
      BoundColumn     =   "F_Codigo"
      Text            =   ""
      Object.DataMember      =   ""
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
   Begin VB.TextBox TXT_OBS 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   5355
      Width           =   4695
   End
   Begin VB.TextBox TXT_FERIAS 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   4080
      Width           =   4695
   End
   Begin VB.ComboBox TXT_MES 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frm_Cad_Fic_Mensal.frx":1D10
      Left            =   2880
      List            =   "frm_Cad_Fic_Mensal.frx":1D38
      TabIndex        =   2
      Top             =   1200
      Width           =   780
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   1482
      ButtonWidth     =   2725
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Fechar/Cancelar"
            Key             =   "fechar"
            Object.ToolTipText     =   "Cancelar e Fechar (Alt+C)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "salvar"
            Object.ToolTipText     =   "Salvar Alteração (Alt+S)"
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4680
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
               Picture         =   "frm_Cad_Fic_Mensal.frx":1D63
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Fic_Mensal.frx":207D
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Fic_Mensal.frx":2397
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Fic_Mensal.frx":26B1
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Fic_Mensal.frx":29CB
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Fic_Mensal.frx":2CE5
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSDataListLib.DataCombo TXT_FUNC_COD 
      Bindings        =   "frm_Cad_Fic_Mensal.frx":2FFF
      Height          =   360
      Left            =   360
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      ListField       =   "F_CODIGO"
      BoundColumn     =   "F_Codigo"
      Text            =   ""
      Object.DataMember      =   ""
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
   Begin MSDataListLib.DataCombo TXT_FUNC_CPF 
      Bindings        =   "frm_Cad_Fic_Mensal.frx":3018
      Height          =   360
      Left            =   2280
      TabIndex        =   8
      Top             =   3240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      ListField       =   "F_CPF"
      BoundColumn     =   "F_Codigo"
      Text            =   ""
      Object.DataMember      =   ""
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
   Begin MSDataListLib.DataCombo TXTLOGO 
      Bindings        =   "frm_Cad_Fic_Mensal.frx":3031
      DataField       =   "F_COD_L"
      DataSource      =   "ADOREG"
      Height          =   420
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   741
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "COD_LOJ"
      BoundColumn     =   "COD_LOJ"
      Text            =   ""
      Object.DataMember      =   "TAB_L"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo TXTLOGO2 
      Bindings        =   "frm_Cad_Fic_Mensal.frx":3042
      DataField       =   "F_COD_L"
      DataSource      =   "ADOREG"
      Height          =   420
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   741
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "NUM"
      BoundColumn     =   "COD_LOJ"
      Text            =   ""
      Object.DataMember      =   "TAB_L_NUM"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label8 
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
      Left            =   360
      TabIndex        =   23
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CÓDIGO"
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
      TabIndex        =   21
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "CPF"
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
      Left            =   2280
      TabIndex        =   22
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   405
      TabIndex        =   20
      Top             =   1650
      Width           =   375
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "®"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   19
      Top             =   1635
      Width           =   435
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "ANOTAÇÃO"
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
      Left            =   390
      TabIndex        =   18
      Top             =   6360
      Width           =   1335
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
      Left            =   3840
      TabIndex        =   17
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label3 
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
      Left            =   360
      TabIndex        =   16
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVAÇÃO"
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
      TabIndex        =   15
      Top             =   5115
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(F)"
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
      TabIndex        =   14
      Top             =   3825
      Width           =   2175
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
      Left            =   3000
      TabIndex        =   13
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   6975
      Left            =   120
      Top             =   840
      Width           =   5175
   End
End
Attribute VB_Name = "frm_Cad_Fic_Mensal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w_logo As String

Private Sub BarraF_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err1
   
    Select Case Button.key
        Case "fechar": Fechar
        Case "salvar": Salvar
        Case "cancelar": Cancelar
    End Select

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


'*** Rotinas ***
Private Sub Cancelar()
On Error GoTo err1
    TXT_FUNC = ""
    TXT_MES = CDbl(Format(Date, "MM"))
    TXT_ANO = Format(Date, "yyyy")
    TXT_FERIAS = ""
    TXT_OBS = ""
    TXT_ANOTACAO = ""
    txt_DT_ADM = ""
    TXT_DT_DEM = ""
    TXT_DT_REG = ""
    txtLogo.BoundText = ""

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Fechar()
On Error GoTo err1
    'If de.rsTAB_FICHA_MENS.State = 1 Then de.rsTAB_FICHA_MENS.Requery
    Unload Me
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Salvar()
Dim w_tipo
On Error GoTo err1

    If txt_DT_ADM = "" Then
        If (MsgBox("Você está cadastrando uma nova ficha mensal SEM data de admissão. Deseja continuar mesmo assim?", vbYesNo, "Data de Admissão em branco")) = vbNo Then
            Exit Sub
        End If
    End If
    
    'If ((IsNull(de.cnc.Execute("SELECT F_DT_DEM FROM TAB_FUNCIONARIO WHERE F_CODIGO = " & TXT_FUNC.BoundText).Fields(0))) And (de.cnc.Execute("SELECT COUNT(M_NFICHA) FROM TAB_FICHA_MENS WHERE M_F_COD = " & TXT_FUNC.BoundText).Fields(0)) > 0) Then
    '    MsgBox "Ainda existem fichas abertas para o funcionário " & UCase(TXT_FUNC) & "!", vbCritical, "Não foi possível criar Nova Ficha"
    'Else
        If TXT_MES <> "" And TXT_ANO <> "" And TXT_FUNC <> "" Then
            
            'If de.cnc.Execute("Select m_bloq FROM tab_ficha_mens Where m_mes = " & TXT_MES & " and m_ano = " & TXT_ANO & " AND M_BLOQ = -1 AND M_DT_ACF = ''").RecordCount = 0 Then
            If de.cnc.Execute("Select m_nficha FROM tab_ficha_mens Where m_mes = " & TXT_MES & " and m_ano = " & TXT_ANO & " and m_DT_DEM = 0 and M_F_COD = " & TXT_FUNC.BoundText).RecordCount = 0 Then
                'txtLogo.BoundText = TXT_FUNC.BoundText
                
                'Ultima Ficha para SALDO DEVEDOR
                Dim ultFicha
                ultFicha = de.cnc.Execute("SELECT MAX(M_NFICHA) From TAB_FICHA_MENS WHERE M_F_COD = " & TXT_FUNC.BoundText).Fields(0)
                
                w_tipo = de.cnc.Execute("SELECT F_TIPO FROM TAB_FUNCIONARIO WHERE F_Codigo = " & TXT_FUNC.BoundText).Fields(0)
                de.cmdIncluirFichaMensal TXT_ANO, TXT_MES, TXT_FUNC.BoundText, TXT_FERIAS, TXT_OBS, txtLogo, TXT_ANOTACAO, TXT_FUNC.Text, w_tipo
                W_NFICHA = de.cnc.Execute("SELECT MAX(M_NFICHA ) FROM TAB_FICHA_MENS").Fields(0)
                
                de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS = '" & TXT_FERIAS & "', F_OBS = '" & TXT_OBS & "', F_ANOTACAO = '" & TXT_ANOTACAO & "' WHERE (F_Codigo = " & TXT_FUNC.BoundText & " )", w_reg
                If txt_DT_ADM <> "" Then
                    de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_Dt_ADM = '" & txt_DT_ADM & "' WHERE (F_Codigo = " & TXT_FUNC.BoundText & " )", w_reg
                    de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_Dt_ADM = '" & txt_DT_ADM & "' WHERE M_NFICHA = " & W_NFICHA & "", w_reg
                End If

                'DT_DEM
                de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_Dt_DEM = NULL WHERE (F_Codigo = " & TXT_FUNC.BoundText & " )", w_reg
                de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_DEM_OK = 0 WHERE (F_Codigo = " & TXT_FUNC.BoundText & " )", w_reg
                de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_Dt_DEM = NULL WHERE M_NFICHA = " & W_NFICHA & "", w_reg
                de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_DEM_OK = 0 WHERE M_NFICHA = " & W_NFICHA & "", w_reg

                If TXT_DT_REG <> "" Then
                    de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_Dt_REG = '" & TXT_DT_REG & "' WHERE (F_Codigo = " & TXT_FUNC.BoundText & " )", w_reg
                    de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_Dt_REG = '" & Format(TXT_DT_REG, "DD/MM/YYYY") & "' WHERE M_NFICHA = " & W_NFICHA & "", w_reg
                End If
                If txtLogo.BoundText <> w_logo Then 'Se alterou a loja
                    de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_COD_L = '" & txtLogo.BoundText & "' WHERE (F_Codigo = " & TXT_FUNC.BoundText & " )", w_reg
                End If
                
              'Salario Família
                w_num_filhos = de.cnc.Execute("SELECT F_NUM_FILHOS FROM TAB_FUNCIONARIO WHERE F_COdigo = " & TXT_FUNC.BoundText).Fields(0)
                w_pg_sal_fam = de.cnc.Execute("SELECT F_PG_SAL_FAM FROM TAB_FUNCIONARIO WHERE F_COdigo = " & TXT_FUNC.BoundText).Fields(0)
                If w_pg_sal_fam Then w_pg_sal_fam = 1 Else w_pg_sal_fam = 0
                de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_NUM_FILHOS = " & w_num_filhos & ", M_PG_SAL_FAM = " & w_pg_sal_fam & " WHERE M_NFICHA = " & W_NFICHA & "", w_reg
                If w_pg_sal_fam = 1 And w_num_filhos > 0 Then
                    Dim wSalFam
                    
                    wSalFam = de.cnc.Execute("Select Sal_Familia from tab_config").Fields(0)
                    
                    wValor = 0
                    wValor = Format(w_num_filhos * wSalFam, "0.00")  'Calcula Salario
                    wDesc = "(" & Format(wSalFam, "0.00") & " x " & w_num_filhos & ") = " & Format(wValor, "0.00")
                    de.cnc.Execute ("DELETE FROM TAB_DESC_CALC Where (C_TP_CONTA = 26) And (C_N_FICHA = " & W_NFICHA & ")")
                    de.cmdIncluirDescCalc Date, W_NFICHA, 26, "+", wValor, wDesc, "", "0", "0", "0", "0"
                End If
                
                
                'Programados (fixos)
                Dim adoFixos As ADODB.Recordset
                
                Set adoFixos = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_EMP_COD = " & TXT_FUNC.BoundText).Clone
             
                Do While Not adoFixos.EOF
                    'Se for diferente de Vale Transporte
                    If adoFixos.Fields("CF_TP_CONTA") <> 109 And adoFixos.Fields("CF_TP_CONTA") <> 110 And adoFixos.Fields("CF_TP_CONTA") <> 111 Then
                        de.cmdIncluirDescCalc2 Date, W_NFICHA, adoFixos.Fields("CF_TP_CONTA"), adoFixos.Fields("CF_TP_OP"), adoFixos.Fields("CF_VALOR"), adoFixos.Fields("CF_DESC"), "0", adoFixos.Fields("CF_CODIGO"), "0", "0", adoFixos.Fields("CF_EMP_COD"), 0
                    End If
                    adoFixos.MoveNext
                Loop
                
                
                'Vale Transporte
                w_pg_vt = de.cnc.Execute("SELECT F_PG_VT FROM TAB_FUNCIONARIO WHERE F_COdigo = " & TXT_FUNC.BoundText).Fields(0)
                de.cnc.Execute "UPDATE TAB_FICHA_MENS SET  M_PG_VT = " & CInt(w_pg_vt) & " WHERE M_NFICHA = " & W_NFICHA & "", w_reg
                'Pagto Vale Transporte
                If w_pg_vt Then
                    Dim fichaAtual As String
                    fichaAtual = W_NFICHA
                
                    Dim adoFixosVT As ADODB.Recordset
                
                    Dim ultimoFixo As String
           
                    de.cmdIncluirDescCalcFixo Now(), TXT_FUNC.BoundText, "109", "-", "0", "INSS 8% do piso [GERADO AUTOMATICAMENTE]"
                    ultimoFixo = de.cnc.Execute("SELECT Max([CF_CODIGO]) FROM TAB_DESC_CALC_FIXO").Fields(0)
                    Set adoFixosVT = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_CODIGO = " & ultimoFixo).Clone
                    de.cmdIncluirDescCalc2 Date, fichaAtual, adoFixosVT.Fields("CF_TP_CONTA"), adoFixosVT.Fields("CF_TP_OP"), adoFixosVT.Fields("CF_VALOR"), adoFixosVT.Fields("CF_DESC"), "0", adoFixosVT.Fields("CF_CODIGO"), "0", "0", adoFixosVT.Fields("CF_EMP_COD"), 0
        
                    ultimoFixo = Empty
                    Set adoFixosVT = Nothing
           
                    de.cmdIncluirDescCalcFixo Now(), TXT_FUNC.BoundText, "110", "-", "0", "Vale Transporte 6% do piso [GERADO AUTOMATICAMENTE]"
                    ultimoFixo = de.cnc.Execute("SELECT Max([CF_CODIGO]) FROM TAB_DESC_CALC_FIXO").Fields(0)
                    Set adoFixosVT = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_CODIGO = " & ultimoFixo).Clone
                    de.cmdIncluirDescCalc2 Date, fichaAtual, adoFixosVT.Fields("CF_TP_CONTA"), adoFixosVT.Fields("CF_TP_OP"), adoFixosVT.Fields("CF_VALOR"), adoFixosVT.Fields("CF_DESC"), "0", adoFixosVT.Fields("CF_CODIGO"), "0", "0", adoFixosVT.Fields("CF_EMP_COD"), 0
           
                    ultimoFixo = Empty
                    Set adoFixosVT = Nothing
           
                    de.cmdIncluirDescCalcFixo Now(), TXT_FUNC.BoundText, "111", "=", "0", "Pagto. de passes (vale transporte) [GERADO AUTOMATICAMENTE]"
                    ultimoFixo = de.cnc.Execute("SELECT Max([CF_CODIGO]) FROM TAB_DESC_CALC_FIXO").Fields(0)
                    Set adoFixosVT = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_CODIGO = " & ultimoFixo).Clone
                    de.cmdIncluirDescCalc2 Date, fichaAtual, adoFixosVT.Fields("CF_TP_CONTA"), adoFixosVT.Fields("CF_TP_OP"), adoFixosVT.Fields("CF_VALOR"), adoFixosVT.Fields("CF_DESC"), "0", adoFixosVT.Fields("CF_CODIGO"), "0", "0", adoFixosVT.Fields("CF_EMP_COD"), 0
           
                    fichaAtual = Empty
                    ultimoFixo = Empty
                    Set adoFixosVT = Nothing
                End If
                
                'SALDO NEGATIVO ANTERIOR
                Dim vrVenda, vrFixo, vrMinimo, percComis, vrSalario, vrComis, sql
                Dim ww_mes, ww_ano, qtdeSaldoAdded
                qtdeSaldoAdded = 0
           
                'ww_mes = TXT_MES - 1
                'If ww_mes = 0 Then
                '    ww_mes = 12
                '    ww_ano = TXT_ANO - 1
                'Else
                '    ww_ano = TXT_ANO
                'End If
                    
                Dim proxFicha
                'ultFicha = de.cnc.Execute("SELECT M_NFICHA From TAB_FICHA_MENS WHERE M_ANO = " & ww_ano & " AND M_MES = " & ww_mes & " AND M_F_COD = " & TXT_FUNC.BoundText).Fields(0)
                If Not (IsNull(ultFicha)) Then
            
                    Dim ADO_TOTAL As ADODB.Recordset
                    Dim wTXT_MAIS
                    Dim wTXT_MENOS
                    Dim wTXT_TOTAL
                      
                    wTXT_MAIS = 0
                    wTXT_MENOS = 0
                    wTXT_TOTAL = 0
                    
                    proxFicha = W_NFICHA
                
                    sql = "SELECT Format([C_DATA_INTERNA],'dd/mm/yy') AS DT_LCTO, TAB_DESC_CALC.C_TP_CONTA AS COD, TAB_TP_CONTA.TP_DESC AS TIPO_CONTA, TAB_DESC_CALC.C_DESC AS DESCRICAO, Format([C_VALOR],'R$ ###,##0.00') AS VALOR, TAB_DESC_CALC.C_TP_OP AS OP, Format([C_VISTO],'Yes/No') AS VISTO, TAB_DESC_CALC.C_CODIGO, TAB_DESC_CALC.C_NCRED FROM TAB_DESC_CALC, TAB_TP_CONTA WHERE (((TAB_DESC_CALC.C_TP_CONTA)=[TAB_TP_CONTA].[TP_COD]) AND ((TAB_DESC_CALC.C_N_FICHA)=" & ultFicha & ")) ORDER BY TAB_DESC_CALC.C_TP_OP, C_DT;"
                
                    Set ADO_TOTAL = de.cnc.Execute(sql).Clone
                        
                    If Not ADO_TOTAL.EOF Then
                        ADO_TOTAL.MoveFirst
                        Do While Not ADO_TOTAL.EOF
                            If ADO_TOTAL.Fields("VALOR") >= 0 And ADO_TOTAL.Fields("OP") = "+" Then
                                wTXT_MAIS = CDbl(wTXT_MAIS) + ADO_TOTAL.Fields("VALOR")
                            ElseIf ADO_TOTAL.Fields("VALOR") < 0 And ADO_TOTAL.Fields("OP") = "-" Then
                                wTXT_MENOS = CDbl(wTXT_MENOS) + ADO_TOTAL.Fields("VALOR")
                            End If
                            ADO_TOTAL.MoveNext
                        Loop
                        
                        wTXT_TOTAL = CDbl(wTXT_MAIS) + CDbl(wTXT_MENOS)
                    End If
                 

                    Dim w_desc
            
                    If wTXT_TOTAL < 0 And txtLogo <> "98" And Not (IsEmpty(ultFicha)) Then
                        w_desc = "Pg. Saldo Dev.: " & Format(wTXT_TOTAL, "R$ 0.00")
                        de.cmdIncluirDescCalcVistado Date, proxFicha, 14, "-", wTXT_TOTAL, w_desc, "", "0", "0", "0", TXT_FUNC.BoundText
                        qtdeSaldoAdded = qtdeSaldoAdded + 1
                    End If
                
                End If
        
                'Lancamentos
                'w_ck_vt = ck_pg_vt
             
                MsgBox "Foi criada com sucesso a ficha do mês " & TXT_MES & "/" & TXT_ANO & " para o funcionário " & TXT_FUNC & ".", vbInformation
                Fechar
            Else
                MsgBox "Você não pode cadastrar 2 fichas para o mesmo funcionário no mesmo mês, sem que a anterior tenha Data de Demissão!", vbExclamation
            End If
            
        Else
            MsgBox "Preencha os Campos!", vbCritical
        End If
    'End If
sair:
    Exit Sub
err1:
    If Err.Number = -2147467259 Then
        MsgBox "Este funcionário já foi cadastrado neste mês e ano!", vbExclamation
    Else
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    End If
    Resume sair
    
End Sub





Private Sub Form_Activate()
    TXT_MES.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
    
    Set adoFunc.Recordset = de.cnc.Execute("Select f_Codigo, f_nome, f_cpf from tab_funcionario").Clone
    adoFunc.Recordset.Sort = "f_nome"
    
    Set adoFuncCod.Recordset = adoFunc.Recordset.Clone
    adoFuncCod.Recordset.Sort = "f_Codigo"
    
    Set adoFuncCPF.Recordset = adoFunc.Recordset.Clone
    adoFuncCPF.Recordset.Sort = "f_cpf"
    
    Cancelar
    TXT_FUNC.BoundText = w_Func_atual
    TXT_FUNC_COD.BoundText = w_Func_atual
    TXT_FUNC_CPF.BoundText = w_Func_atual
    
    w_logo = de.cnc.Execute("Select f_cod_L from tab_funcionario where F_Codigo = " & w_Func_atual).Fields(0)
    txtLogo.BoundText = w_logo
    txtLogo2.BoundText = w_logo
    
End Sub

Private Sub TXT_ANO_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub TXT_ANO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub TXT_ANO_Validate(Cancel As Boolean)
    If Not (TXT_ANO >= 1990 And TXT_ANO <= 2200) Then
        MsgBox "Você deve digitar o Nº Ano com 4 digitos!", vbInformation
        TXT_ANO.SetFocus
    End If
End Sub



Private Sub txt_ANOTACAO_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{BACKSPACE}"
        SendKeys "{tab}"
        Pause 0.3
        If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1) Then Salvar
      End If
End Sub



Private Sub txt_DT_ADM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{tab}"
End Sub



Private Sub TXT_DT_DEM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub TXT_DT_REG_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub TXT_FERIAS_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{BACKSPACE}"
        SendKeys "{tab}"
      End If
End Sub

Private Sub TXT_FUNC_Change()
On Error Resume Next
   Dim adoF As ADODB.Recordset
   
   Set adoF = de.cnc.Execute("Select F_Ferias, F_Obs, F_Anotacao, F_Dt_ADM, F_Dt_DEM, F_Dt_Reg  from tab_Funcionario Where F_Codigo = " & TXT_FUNC.BoundText & "").Clone
   TXT_ANOTACAO = IIf(IsNull(adoF.Fields("F_Anotacao")), "", adoF.Fields("F_Anotacao"))
   TXT_FERIAS = IIf(IsNull(adoF.Fields("F_Ferias")), "", adoF.Fields("F_Ferias"))
   TXT_OBS = IIf(IsNull(adoF.Fields("F_Obs")), "", adoF.Fields("F_Obs"))
    
   'txt_DT_ADM = IIf(IsNull(adoF.Fields("F_Dt_ADM")), "", adoF.Fields("F_Dt_ADM"))
   'TXT_DT_DEM = IIf(IsNull(adoF.Fields("F_Dt_DEM")), "", adoF.Fields("F_Dt_DEM"))
   'TXT_DT_REG = IIf(IsNull(adoF.Fields("F_Dt_REG")), "", adoF.Fields("F_Dt_REG"))

End Sub

Private Sub TXT_FUNC_Click(Area As Integer)
    If TXT_FUNC.BoundText <> "" Then
        TXT_FUNC_COD.BoundText = TXT_FUNC.BoundText
        TXT_FUNC_CPF.BoundText = TXT_FUNC.BoundText
    End If
End Sub

Private Sub TXT_FUNC_COD_Click(Area As Integer)
    If TXT_FUNC_COD.BoundText <> "" Then
        TXT_FUNC.BoundText = TXT_FUNC_COD.BoundText
        TXT_FUNC_CPF.BoundText = TXT_FUNC_COD.BoundText
    End If
End Sub

Private Sub TXT_FUNC_CPF_Click(Area As Integer)
    If TXT_FUNC_CPF.BoundText <> "" Then
        TXT_FUNC_COD.BoundText = TXT_FUNC_CPF.BoundText
        TXT_FUNC.BoundText = TXT_FUNC_CPF.BoundText
    End If
End Sub

Private Sub TXT_FUNC_GotFocus()
 SendKeys "{F4}"
End Sub

Private Sub TXT_FUNC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub TXT_LOGO_Click(Area As Integer)

End Sub

Private Sub TXT_LOGO2_Click(Area As Integer)

End Sub

Private Sub TXT_MES_GotFocus()
    SendKeys "{f4}"
End Sub

Private Sub TXT_MES_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub TXT_MES_Validate(Cancel As Boolean)
    If Not (CDbl(IIf(Not (IsNumeric(TXT_MES)), 0, TXT_MES)) >= 1 And CDbl(IIf(Not (IsNumeric(TXT_MES)), 0, TXT_MES)) <= 12) Then
        MsgBox "Você deve digitar o Nº Mês!", vbInformation
        TXT_MES.SetFocus
    End If
End Sub

'--------- Ao Pressionar uma Tecla -----------
Private Sub txt_ANOTACAO_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = Keys(KeyCode, Shift)
End Sub

Private Sub TXT_OBS_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift <> 2 And KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{BACKSPACE}"
        SendKeys "{tab}"
      End If
End Sub

Private Sub txt_OBS_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = Keys(KeyCode, Shift)
End Sub
Private Sub txt_FERIAS_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = Keys(KeyCode, Shift)
End Sub
Private Sub TXT_FUNC_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = Keys(KeyCode, Shift)
End Sub
Private Sub TXT_mes_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = Keys(KeyCode, Shift)
End Sub
Private Sub TXT_ano_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = Keys(KeyCode, Shift)
End Sub

Private Sub GRID_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = Keys(KeyCode, Shift)
End Sub




' -------  Teclas de Atalhos --------
Function Keys(KeyCode As Integer, Shift As Integer) As Integer
    '*** Shift (4 = Alt) ***
    If Shift = 4 Then
        Select Case KeyCode
        Case 70: ' "F"
                Fechar
        Case 83: ' "S"
                Salvar
        Case 67: ' "C"
                Cancelar
        End Select
    End If
End Function

Private Sub txtLogo_Click(Area As Integer)
    txtLogo2.BoundText = txtLogo.BoundText
End Sub

Private Sub txtLogo2_Click(Area As Integer)
    txtLogo.BoundText = txtLogo2.BoundText
End Sub
