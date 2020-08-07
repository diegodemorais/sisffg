VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "msCOMCTL.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form frm_Vendas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamentos"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   Icon            =   "frm_Alt_Gerente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5160
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   8454143
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   8454143
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Gerente.frx":1042
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Gerente.frx":135C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Gerente.frx":1676
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Gerente.frx":1990
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Gerente.frx":1CAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Gerente.frx":1FC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Gerente.frx":22DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Gerente.frx":2BB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Gerente.frx":48C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Gerente.frx":4BDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Gerente.frx":502E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Gerente.frx":5348
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Gerente.frx":566A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Gerente.frx":7E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Gerente.frx":826E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox ckNovo 
      Caption         =   "Novo"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox TXT_CODIGO 
      Alignment       =   2  'Center
      DataField       =   "V_VR"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox TXT_DATA 
      Alignment       =   2  'Center
      DataField       =   "V_DATA"
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
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc adoReg 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   7005
      Width           =   8130
      _ExtentX        =   14340
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
      Caption         =   "Registro(s):"
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
   Begin VB.TextBox TXT_LOJA 
      Alignment       =   2  'Center
      DataField       =   "V_F_LOJA"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid Grid 
      Bindings        =   "frm_Alt_Gerente.frx":AA20
      Height          =   6135
      Left            =   3840
      TabIndex        =   3
      Top             =   840
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   10821
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAcrossSplits =   -1  'True
      TabAction       =   2
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "LANÇAMENTOS"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "V_F_LOJA"
         Caption         =   "B"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "V_DATA"
         Caption         =   "DATA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "V_VR"
         Caption         =   "COD"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   1482
      ButtonWidth     =   1667
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Fechar"
            Key             =   "fechar"
            Object.ToolTipText     =   "Fechar (Alt+F)"
            ImageIndex      =   1
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Novo"
            Key             =   "novo"
            Object.ToolTipText     =   "Novo Registro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Editar"
            Key             =   "editar"
            Object.ToolTipText     =   "Editar Alteração (Alt+E)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Salvar"
            Key             =   "salvar"
            Object.ToolTipText     =   "Salvar Alteração (Alt+S)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Cancelar"
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar Alteração (Alt+C)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E&xcluir"
            Key             =   "excluir"
            Object.ToolTipText     =   "Excluir registro (Alt+X)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fil&trar"
            Key             =   "filtrar"
            Object.ToolTipText     =   "Filtrar (Alt+T)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gerar"
            Key             =   "automatico"
            Object.ToolTipText     =   "Gerar Automático (mwts)"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
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
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
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
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label2 
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
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   6135
      Left            =   120
      Top             =   840
      Width           =   3615
   End
End
Attribute VB_Name = "frm_Vendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W_LD_FILTRO As Boolean


Private Sub Form_Load()
    
On Error Resume Next

On Error GoTo err1
    
    
    If de.rsTab_Venda.State = 1 Then de.rsTab_Venda.Close
    de.TAB_VENDA
  
    
    
sair:
    
    Set adoReg.Recordset = de.rsTab_Venda.Clone  'de.cnc.Execute("select * from tab_funcionario order by f_nome")
    adoReg.Recordset.Sort = "V_DATA"
    
    
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

'*** Caption no navegador ***
Private Sub adoReg_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1

If Not IsNumeric(adoReg.Recordset.RecordCount) Then adoReg.Caption = "REGISTRO : " & adoReg.Recordset.AbsolutePosition & " / " & adoReg.Recordset.RecordCount & IIf(W_LD_FILTRO = True, " (FILTRADO)", "")

sair:
    Exit Sub
err1:
    If Not Err.Number = -2147217885 And Not Err.Number = -2147467259 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


'** Barra de Ferramenta ***
Private Sub BarraF_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
        Case "fechar": Fechar
        Case "editar": Editar
        Case "salvar": Salvar
        Case "cancelar": Cancelar
        Case "excluir": Excluir
        Case "filtrar": FILTRAR
        Case "novo": Novo
        Case "automatico": Automatico
    End Select
End Sub


'*** Rotinas ***
Private Sub Cancelar()
On Error GoTo err1

    pos = adoReg.Recordset.AbsolutePosition - 1
    adoReg.Recordset.CancelBatch adAffectCurrent
    W_FILTRO = adoReg.Recordset.Filter
    adoReg.Refresh
    adoReg.Recordset.Filter = W_FILTRO
    Editar
    adoReg.Recordset.Sort = "V_DATA"
    adoReg.Recordset.Move pos


    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Automatico()

Dim dtIni, dtFim As Date
        
    frm_ESCOLHA_DATA.Show 1
    
    dtIni = CDate(frm_ESCOLHA_DATA.TXT_DT_INICIAL)
    
    dtFim = CDate(frm_ESCOLHA_DATA.TXT_DT_FINAL)
   
     
    If de.rscmdCod.State = 1 Then de.rscmdCod.Close
    
    de.cmdCod dtIni, dtFim
    
    Do While Not de.rscmdCod.EOF
        de.cnc.Execute ("INSERT INTO TAB_VENDA(V_F_LOJA,V_DATA,V_VR) VALUES ('" & Format(de.rscmdCod.Fields("B"), "000") & "','" & dtIni & "','" & Format(de.rscmdCod.Fields("COD"), "0.000") & "');")
        
        de.rscmdCod.MoveNext
    Loop
    
    '99
    de.cnc.Execute ("INSERT INTO TAB_VENDA(V_F_LOJA,V_DATA,V_VR) VALUES ('" & Format("999", "000") & "','" & dtIni & "','" & Format("0,01", "0.000") & "');")
    de.cnc.Execute ("INSERT INTO TAB_VENDA(V_F_LOJA,V_DATA,V_VR) VALUES ('" & Format("099", "000") & "','" & dtIni & "','" & Format("0,01", "0.000") & "');")
    
    MsgBox "Inclusão automática de Códigos feito com sucesso na data de: " & dtIni & " !", vbInformation, "Inclusão Automática de Códigos"
    
    
        
End Sub


Private Sub Novo()
On Error GoTo err1
If Not adoReg.Recordset.EOF Then
    BarraF.Buttons("salvar").Enabled = Not BarraF.Buttons("salvar").Enabled
    BarraF.Buttons("cancelar").Enabled = Not BarraF.Buttons("cancelar").Enabled
    BarraF.Buttons("editar").Enabled = Not BarraF.Buttons("editar").Enabled
    BarraF.Buttons("novo").Enabled = Not BarraF.Buttons("novo").Enabled
    
    Grid.Enabled = Not Grid.Enabled
    
    TXT_LOJA.DataField = ""
    TXT_DATA.DataField = ""
    TXT_CODIGO.DataField = ""
    
    TXT_LOJA.text = ""
    TXT_DATA.text = ""
    TXT_CODIGO.text = ""
    TXT_LOJA.Enabled = True
    TXT_DATA.Enabled = True
    TXT_CODIGO.Enabled = True
    
    ckNovo.value = Checked
       
Else
    MsgBox "Não existe registro para editar!", vbExclamation
End If
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub Editar()
On Error GoTo err1
ckNovo.value = Unchecked

If Not adoReg.Recordset.EOF Then
    BarraF.Buttons("salvar").Enabled = Not BarraF.Buttons("salvar").Enabled
    BarraF.Buttons("cancelar").Enabled = Not BarraF.Buttons("cancelar").Enabled
    BarraF.Buttons("editar").Enabled = Not BarraF.Buttons("editar").Enabled
    BarraF.Buttons("novo").Enabled = Not BarraF.Buttons("novo").Enabled
    
    Grid.Enabled = Not Grid.Enabled
    'TXT_LOJA.Enabled = Not TXT_LOJA.Enabled
    'TXT_DATA.Enabled = Not TXT_DATA.Enabled
    TXT_CODIGO.Enabled = Not TXT_CODIGO.Enabled
       
Else
    MsgBox "Não existe registro para editar!", vbExclamation
End If
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Excluir()
On Error GoTo err1
  
If Not adoReg.Recordset.EOF Then
    
    If vbYes = MsgBox("DESEJA REALMENTE EXCLUIR O LANÇAMENTO (" & TXT_LOJA & " - " & TXT_DATA & ")?", vbQuestion + vbYesNo) Then
        adoReg.Recordset.Delete
        adoReg.Recordset.UpdateBatch

    End If
   
 Else
    MsgBox "Não existe registro para excluir!", vbExclamation
End If
   
sair:
    Exit Sub
err1:
    If Not Err.Number = -2147467259 Then
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    Else
        MsgBox "NÃO É POSSÍVEL EXCLUIR ESTE LANÇAMENTO! (RELACIONAMENTOS?)", vbCritical
        adoReg.Refresh
    End If
    Resume sair
End Sub

Private Sub Fechar()
On Error GoTo err1
    de.rsTab_Venda.Requery
    de.rsTab_Venda.Close
sair:
    Unload Me
    Exit Sub
err1:
'    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub FILTRAR()
Dim w_resp As String
Dim W_CAMPO As String
Dim W_FILTRO As String

On Error GoTo err1
    
    'w_resp = InputBox("FILTRAR PELO QUÊ ? ENTRE COM O NÚMERO DA OPÇÃO DESEJADA." & Chr(13) & Chr(13) & "1 - NOME" & Chr(13) & "2 - LOGO" & Chr(13) & "3 - DATA ADMISSÃO" & Chr(13) & "4 - DATA DE REGISTRO" & Chr(13) & "5 - DATA DE DEMISSÃO" & Chr(13) & "6 - ADMITIDOS" & Chr(13) & "7 - REMOVER FILTRO *", , "1")
    w_resp = InputBox("FILTRAR PELO QUÊ ? ENTRE COM O NÚMERO DA OPÇÃO DESEJADA." & Chr(13) & Chr(13) & "1 - (B)" & Chr(13) & "2 - DATA" & Chr(13) & "3 - REMOVER FILTRO *", , "1")
    
    If Not w_resp = "" And IsNumeric(w_resp) And w_resp >= 1 And w_resp <= 3 Then
        Select Case w_resp
        'NOME
        Case 1:
            w_resp = "(B)"
            W_CAMPO = "V_F_LOJA"
        'LOGO
        Case 2:
            w_resp = "DATA"
            W_CAMPO = "V_DATA"
        'DT_ADM
      
        Case 3:
            If Not adoReg.Recordset.Filter = 0 Then
                W_LD_FILTRO = False
                adoReg.Recordset.Filter = 0
                adoReg.Refresh
            End If
        End Select
        If Not w_resp = "3" Then
            
            If w_resp = "DATA" Then
                frm_ESCOLHA_DATA.Show 1
                W_FILTRO = W_CAMPO & " >= #" & frm_ESCOLHA_DATA.TXT_DT_INICIAL & "# AND " & W_CAMPO & " <= #" & frm_ESCOLHA_DATA.TXT_DT_FINAL & "#"
                W_LD_FILTRO = True
                adoReg.Recordset.Filter = W_FILTRO
            Else
                W_FILTRO = InputBox("ENTRE COM A " & w_resp & " DESEJADO!")
                W_FILTRO = W_CAMPO & " LIKE '%" & W_FILTRO & "%'"
                W_LD_FILTRO = True
                adoReg.Recordset.Filter = W_FILTRO
            End If
        End If
    End If
    
sair:
    Exit Sub
err1:
    If Err.Number <> 13 And Err.Number <> 3265 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
        W_LD_FILTRO = False
        Resume sair

End Sub

Private Sub Salvar()
On Error GoTo err1
   
adoReg.Recordset.UpdateBatch adAffectCurrent
   
   If ckNovo.value = Checked Then
   
    de.cnc.Execute "INSERT INTO TAB_VENDA (V_F_LOJA, V_DATA, V_VR) VALUES" & _
                   "('" & Format(TXT_LOJA, "000") & "','" & TXT_DATA & "','" & TXT_CODIGO & "')", w_reg
                   
    TXT_LOJA.DataField = "V_F_LOJA"
    TXT_DATA.DataField = "V_DATA"
    TXT_CODIGO.DataField = "V_VR"
      
   Else
      
     de.cnc.Execute "UPDATE TAB_VENDA SET V_VR = '" & TXT_CODIGO & "'" & _
        " WHERE (V_F_LOJA = " & adoReg.Recordset.Fields("V_F_LOJA") & _
        " AND V_DATA = " & adoReg.Recordset.Fields("V_DATA") & ")", w_reg

   End If
   
   
   de.rsTab_Venda.Requery
   'de.rsTab_Venda.Resync
   adoReg.Refresh
   Cancelar
  
  
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
    
End Sub







'--------- Ao Pressionar uma Tecla -----------

Private Sub ck_pg_SFam_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub








' -------  Teclas de Atalhos --------

Sub Keys(KeyCode As Integer, Shift As Integer)
'*** Shift (4 = Alt) ***
If Shift = 4 Then
    Select Case KeyCode
    Case 70: ' "F"
            Fechar
     Case 69: ' "E"
           If BarraF.Buttons("editar").Enabled = True Then Editar
    Case 83: ' "S"
           If BarraF.Buttons("salvar").Enabled = True Then Salvar
    Case 67: ' "C"
           If BarraF.Buttons("cancelar").Enabled = True Then Cancelar
    Case 88: ' "X"
            Excluir
    Case 84: ' "T"
            FILTRAR
    End Select
End If
End Sub



Private Sub TXT_CODIGO_GotFocus()
Sendkeys "{home}+{end}"
End Sub



Private Sub TXT_DATA_GotFocus()
Sendkeys "{home}+{end}"
End Sub



Private Sub TXT_LOJA_GotFocus()
Sendkeys "{home}+{end}"
End Sub
