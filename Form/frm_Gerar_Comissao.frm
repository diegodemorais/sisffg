VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_Gerar_Comissao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerar Comissões"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   Icon            =   "frm_Gerar_Comissao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ck_Logo 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1320
      Width           =   975
   End
   Begin VB.CheckBox ckRefazer 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Refazer Prêmio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1200
      Width           =   990
   End
   Begin MSAdodcLib.Adodc ADO_FUNC 
      Height          =   375
      Left            =   1440
      Top             =   2040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   5
      CommandTimeout  =   10
      CursorType      =   2
      LockType        =   1
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
   Begin VB.CheckBox ck_Nome 
      BackColor       =   &H00E0E0E0&
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
      Left            =   4320
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2085
      Value           =   1  'Checked
      Width           =   975
   End
   Begin MSComctlLib.StatusBar BStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Top             =   2625
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9975
         EndProperty
      EndProperty
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
      ItemData        =   "frm_Gerar_Comissao.frx":08CA
      Left            =   480
      List            =   "frm_Gerar_Comissao.frx":08F2
      TabIndex        =   0
      Top             =   1200
      Width           =   780
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
      Left            =   1245
      TabIndex        =   1
      Top             =   1200
      Width           =   810
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   1482
      ButtonWidth     =   1376
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Fechar"
            Key             =   "fechar"
            Object.ToolTipText     =   "Fechar (Alt+F)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gerar"
            Key             =   "gerar"
            Object.ToolTipText     =   "Gerar Fichas (Alt + G)"
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1680
         Top             =   -120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Gerar_Comissao.frx":091D
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Gerar_Comissao.frx":0C37
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Gerar_Comissao.frx":0F51
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Gerar_Comissao.frx":126B
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSDataListLib.DataCombo TXT_LOGO 
      Bindings        =   "frm_Gerar_Comissao.frx":1B45
      DataField       =   "F_COD_L"
      DataSource      =   "ADOREG"
      Height          =   405
      Left            =   2520
      TabIndex        =   5
      Top             =   1200
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   714
      _Version        =   393216
      MatchEntry      =   -1  'True
      ListField       =   "COD_LOJ"
      BoundColumn     =   "COD_LOJ"
      Text            =   "%"
      Object.DataMember      =   "TAB_L"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dbNome 
      Bindings        =   "frm_Gerar_Comissao.frx":1B56
      Height          =   315
      Left            =   480
      TabIndex        =   9
      Top             =   2040
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      ListField       =   "F_NOME"
      BoundColumn     =   "F_Codigo"
      Text            =   "%"
      Object.DataMember      =   ""
   End
   Begin VB.Label Label1 
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
      Left            =   480
      TabIndex        =   10
      Top             =   1800
      Width           =   855
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
      Left            =   2520
      TabIndex        =   6
      Top             =   960
      Width           =   615
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
      Left            =   600
      TabIndex        =   4
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
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   120
      Top             =   840
      Width           =   5415
   End
End
Attribute VB_Name = "frm_Gerar_Comissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BarraF_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err1
   
    Select Case Button.key
        Case "fechar": Fechar
        Case "gerar":
        If vbYes = MsgBox("Antes de gerar as comissões, deve ser feito, importações das seguintes tabelas ( Emp., Vendas )!" & Chr(13) & "Deseja Gera assim mesmo?", vbExclamation + vbYesNo) Then Gerar
    End Select

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub




'*** Rotinas ***

Private Sub Fechar()
On Error GoTo err1
    'If de.rsTAB_FICHA_MENS.State = 1 Then de.rsTAB_FICHA_MENS.Requery
    frm_Alt_Fic_Mensal_VIS.Timer1 = True
sair:
    Unload Me
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub ROT_PREMIOS()
Dim ADO_PREMIO As ADODB.Recordset
Dim ADO_P As ADODB.Recordset
Dim ADO_ITENS As ADODB.Recordset

On Error Resume Next


    
'*** Verifica se a tabela pro mês desejado já esta gerada ***
If ckRefazer.value = 1 Then
    '*** LIMPA TABELA DE PREMIOS E INSERI NOVO REGISTRO ***
    BStatus.Panels(1) = "Excluindo Tabela Temporária"
    de.cncDBase.Execute "DELETE FROM Tab_TEMP;"   '*** EXCLUI DADOS DA TABELA ***
    BStatus.Panels(1) = "Gerando Tabela Temp. p/ Cálculos de Prêmios"
    TXT_LOGO = IIf(TXT_LOGO = "", "%", TXT_LOGO)
    de.Criar_Tab_Premios Format(TXT_MES, "00") & "/" & TXT_ANO, TXT_LOGO       '*** TAB_TEMP ***
    '*** DELETA REGISTRO COM DATA E LOJA Q/ NÃO FOREM USAR ***
    BStatus.Panels(1) = "Excluindo registros inúteis"
End If

  
On Error GoTo err1
    
If de.cncDBase.Execute("SELECT * FROM TAB_TEMP WHERE P_LOJA LIKE '" & TXT_LOGO & "'").RecordCount > 0 Then
    '*** SQL ITENS ***
    Set ADO_ITENS = de.cnc.Execute("SELECT LOJA, COD_TAB FROM LOJB135 WHERE LOJA LIKE '" & TXT_LOGO & "' ORDER BY LOJA, COD_TAB")
            
    '*** LOOPING   ENTRE   ITENS E LOJAS
    Do While Not ADO_ITENS.EOF
        BStatus.Panels(1) = " Classificando Prêmios (B):" & ADO_ITENS.Fields("LOJA") & " : Item:" & ADO_ITENS.Fields("COD_TAB")
        'SQL DOS PREMIO  REFERENTE A LOJA E ITENS RELATIVOS A ADO_ITENS CORRENTE
        Set ADO_PREMIO = de.cnc.Execute("SELECT * FROM TAB_TEMP WHERE P_LOJA = '" & ADO_ITENS.Fields("LOJA") & "' AND P_COD_PREMIO = '" & ADO_ITENS.Fields("COD_TAB") & "' ORDER BY P_QTDE DESC").Clone
        
            W_ORDEM = 1
            If Not ADO_PREMIO.EOF Then
                If ADO_PREMIO.Fields("P_QTDE") >= ADO_PREMIO.Fields("P_QT_Min") Then
                            '*** pEGA OS REGISTRO Q/ SÃO IGUAIS  AS QTDE ***
                            'p/ verificar a qtde de vendedores na mesma colocação
                            Set ADO_P = de.cnc.Execute("SELECT P_QTDE FROM TAB_TEMP WHERE P_LOJA = '" & ADO_ITENS.Fields("LOJA") & "' AND P_COD_PREMIO = '" & ADO_ITENS.Fields("COD_TAB") & "' AND P_QTDE = " & ADO_PREMIO.Fields("P_QTDE") & " ORDER BY P_QTDE DESC").Clone
                            
                            'APENAS 1 VEND.  EM 1º COLOCADO
                            If ADO_P.RecordCount = 1 Then     '*** SOMENTE 1 VEND.   EM 1º
                                    
                                    de.cnc.Execute "UPDATE TAB_TEMP SET P_ORDEM = " & W_ORDEM & ", P_VALOR_PG = '" & ADO_PREMIO.Fields("P_PREMIO1") & "', P_QTDE_CLASS = " & ADO_P.RecordCount & " WHERE (P_LOJA = '" & ADO_ITENS.Fields("LOJA") & "') AND (P_COD_PREMIO = '" & ADO_ITENS.Fields("COD_TAB") & "') AND (P_QTDE = " & ADO_PREMIO.Fields("P_QTDE") & ")", w_RefAf
                                    If w_RefAf <> ADO_P.RecordCount Then MsgBox "Qtde de Registros salvos diferentes do 1º Colocado!", vbCritical
                                    ADO_PREMIO.MoveNext
                                    W_ORDEM = 2
                                    
                                    If Not ADO_PREMIO.EOF Then
                                        If ADO_PREMIO.Fields("P_QTDE") >= ADO_PREMIO.Fields("P_QT_Min") Then
                                            '*** pEGA OS REGISTRO Q/ SÃO IGUAIS  AS QTDE ***
                                            Set ADO_P = de.cnc.Execute("SELECT P_QTDE FROM TAB_TEMP WHERE P_LOJA = '" & ADO_ITENS.Fields("LOJA") & "' AND P_COD_PREMIO = '" & ADO_ITENS.Fields("COD_TAB") & "' AND P_QTDE = " & ADO_PREMIO.Fields("P_QTDE") & " ORDER BY P_QTDE DESC").Clone
                                            
                                            '2º COLOCADO
                                            de.cnc.Execute "UPDATE TAB_TEMP SET P_ORDEM = " & W_ORDEM & ", P_VALOR_PG = '" & ADO_PREMIO.Fields("P_PREMIO2") / ADO_P.RecordCount & "', P_QTDE_CLASS = " & ADO_P.RecordCount & " WHERE (P_LOJA = '" & ADO_ITENS.Fields("LOJA") & "') AND (P_COD_PREMIO = '" & ADO_ITENS.Fields("COD_TAB") & "') AND (P_QTDE = " & ADO_PREMIO.Fields("P_QTDE") & ")", w_RefAf
                                            If w_RefAf <> ADO_P.RecordCount Then MsgBox "Qtde de Registros salvos diferentes do 2º Colocado!", vbCritical
                                        End If
                                    End If
                
                            
                            'MAIS Q/ 1 VEND.   EM 1º
                            ElseIf ADO_P.RecordCount > 1 And Not ADO_PREMIO.EOF And ADO_PREMIO.Fields("P_QTDE") >= ADO_PREMIO.Fields("P_QT_Min") Then
                            
                                de.cnc.Execute "UPDATE TAB_TEMP SET P_ORDEM = " & W_ORDEM & ", P_VALOR_PG = '" & (ADO_PREMIO.Fields("P_PREMIO1") + ADO_PREMIO.Fields("P_PREMIO2")) / ADO_P.RecordCount & "', P_QTDE_CLASS = " & ADO_P.RecordCount & " WHERE (P_LOJA = '" & ADO_ITENS.Fields("LOJA") & "') AND (P_COD_PREMIO = '" & ADO_ITENS.Fields("COD_TAB") & "') AND (P_QTDE = " & ADO_PREMIO.Fields("P_QTDE") & ")", w_RefAf
                                If w_RefAf <> ADO_P.RecordCount Then MsgBox "Qtde de Registros salvos diferentes dos 1º Colocados!", vbCritical
                            
                            End If
                End If
          End If
        
        ADO_ITENS.MoveNext
    Loop
    
    '*** Excluir os Vendedores q/ tiveram premio ****
    'de.cncDBase.Execute "DELETE * FROM TAB_TEMP WHERE P_VALOR_PG = 0 AND P_ORDEM = 0"
Else
    MsgBox "NÃO EXISTEM VENDAS NESTE PERÍODO PARA GERARMOS OS PRÊMIOS DOS VENDEDORES!", vbCritical
End If
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub



Private Sub Gerar()

On Error Resume Next

    de.rscmdSqlComissao.Resync

On Error GoTo err1
Dim W_ADO_FICHA As ADODB.Recordset
Dim w_ado_Premio As ADODB.Recordset

'*** Rotina pra geração dos premios ***
'ROT_PREMIOS

'*** CONFIRMA GERAÇÃO DE MESES DIFERENTES AO ATUAL ***
V_RESP = vbYes

'*** VERIFICA AS DATAS E SE É DA EPOCA CORRENTE , SENÃO FAZ UMA PERGUNTA***
If TXT_ANO <> Format(Date, "YYYY") Or TXT_MES <> CDbl(Format(Date, "MM")) Then V_RESP = MsgBox("DESEJA REALMENTE GERAR AS COMISSÕES DE UM MÊS QUE NÃO SEJA O ATUAL?", vbQuestion + vbYesNo + vbDefaultButton2)

'*** Verifica se existe COMISSÃO CADASTRADA SENÃO CADASTRA***
'If Not IsNumeric(dbNome.BoundText) Then Exit Sub
    strSQL = "SELECT M_MES, M_ANO FROM TAB_FICHA_MENS WHERE (M_MES = " & CDbl(TXT_MES) & ") AND M_ANO = " & CDbl(TXT_ANO) & " and M_Logo Like '" & TXT_LOGO & "'" & IIf(ck_Nome.value = 1, "", " AND M_F_COD = " & dbNome.BoundText & " AND M_COMISSAO = 'N'")
If V_RESP = vbYes And de.cnc.Execute(strSQL).RecordCount > 0 Then
    
        If TXT_ANO <> "" And TXT_MES <> "" Then
            
            '***  VERIFICA SE EXISTE FICHAS GERADAS NESTE PERIODO***
            If V_RESP = vbYes And de.cnc.Execute("SELECT M_MES, M_ANO FROM TAB_FICHA_MENS WHERE (M_MES = " & CDbl(TXT_MES) & ") AND M_ANO = " & CDbl(TXT_ANO) & " and M_Logo Like '" & TXT_LOGO & "%'" & IIf(ck_Nome.value = 1, "", " AND M_F_COD = " & dbNome.BoundText & "")).RecordCount > 0 Then
                    w_reg = 0
                    
                    '*** ABRE AS COMISSÕES ***
                    If de.rscmdSqlComissao.State = 1 Then de.rscmdSqlComissao.Close
                    de.cmdSqlComissao TXT_LOGO, Format(TXT_MES, "00"), TXT_ANO
                    
                    '*** SQL de Premio ***
                    Set w_ado_Premio = de.cnc.Execute("SELECT P_LOJA, SUM(P_VALOR_PG) AS premio, P_VENDEDOR FROM TAB_temp WHERE P_ORDEM > 0 GROUP BY P_LOJA, P_VENDEDOR").Clone
                    
                    '*** ABRE AS FICHAS ***
'                    Set W_ADO_FICHA = de.cnc.Execute("SELECT TAB_FICHA_MENS.M_MES, TAB_FICHA_MENS.M_ANO, TAB_FICHA_MENS.M_NFICHA, TAB_FICHA_MENS.M_LOGO, MID(TAB_FUNCIONARIO.F_COD_CENTRAL, 3) AS COD_FUNC_CENTRAL, TAB_FUNCIONARIO.F_TIPO as TIPO ,  TAB_FUNCIONARIO.F_VPISO as VPISO, TAB_FUNCIONARIO.F_VPISO_R as VPISO_R, TAB_FUNCIONARIO.F_CX_QT_VND as CX_QT_VND, TAB_FICHA_MENS.M_DT_REG AS DT_REG, TAB_FICHA_MENS.M_DT_DEM AS DT_DEM, TAB_FICHA_MENS.M_DT_ADM AS DT_ADM, TAB_FICHA_MENS.M_F_COD   FROM TAB_FICHA_MENS, TAB_FUNCIONARIO WHERE TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_LOGO LIKE '" & UCase(TXT_LOGO) & "' AND TAB_FICHA_MENS.M_ACORDO = 0 AND (TAB_FICHA_MENS.M_BLOQ = 0)) AND (M_COMISSAO = 'N') ORDER BY TAB_FICHA_MENS.M_LOGO, MID(TAB_FUNCIONARIO.F_COD_CENTRAL, 3)").Clone
                     Set W_ADO_FICHA = de.cnc.Execute("SELECT TAB_FICHA_MENS.M_MES, TAB_FICHA_MENS.M_ANO, TAB_FICHA_MENS.M_NFICHA, TAB_FICHA_MENS.M_LOGO, MID(TAB_FUNCIONARIO.F_COD_CENTRAL, 3) AS COD_FUNC_CENTRAL, TAB_FUNCIONARIO.F_TIPO as TIPO ,  TAB_FUNCIONARIO.F_VPISO as VPISO, TAB_FUNCIONARIO.F_VPISO_R as VPISO_R, TAB_FUNCIONARIO.F_CX_QT_VND as CX_QT_VND, TAB_FICHA_MENS.M_DT_REG AS DT_REG, TAB_FICHA_MENS.M_DT_DEM AS DT_DEM, TAB_FICHA_MENS.M_DT_ADM AS DT_ADM, TAB_FICHA_MENS.M_F_COD   FROM TAB_FICHA_MENS, TAB_FUNCIONARIO WHERE TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ") AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_LOGO LIKE '" & UCase(TXT_LOGO) & "' AND TAB_FICHA_MENS.M_ACORDO = 0 AND (TAB_FICHA_MENS.M_BLOQ = 0))  ORDER BY TAB_FICHA_MENS.M_LOGO, MID(TAB_FUNCIONARIO.F_COD_CENTRAL, 3)").Clone
                   
                    W_ADO_FICHA.Filter = "TIPO <> 'C'"
                    
                    If ck_Nome.value = 0 Then W_ADO_FICHA.Filter = "TIPO <> 'C' AND M_F_COD = " & dbNome.BoundText & ""
                    
                    '*** LOOPING  DAS   FICHAS ***
                    Do While Not W_ADO_FICHA.EOF
                            BStatus.Panels(1) = "Incluindo Contas: " & W_ADO_FICHA.Fields(3)
                            
                            
                            '*** INCLUI AS COMISSÕES ***
                            de.rscmdSqlComissao.Filter = "LOJA = '" & W_ADO_FICHA.Fields(3) & "' AND MES = '" & Format(TXT_MES, "00") & "' AND ANO = '" & TXT_ANO & "' and Vendedor = '" & Format(W_ADO_FICHA.Fields("COD_FUNC_CENTRAL"), "00") & "'"
                            
                            If Not W_ADO_FICHA.Fields("COD_FUNC_CENTRAL") = "" Then
                                '*** INCLUI OS PRÊMIOS ***
                                w_ado_Premio.Filter = "P_LOJA = '" & W_ADO_FICHA.Fields(3) & "' AND P_Vendedor = '" & Format(W_ADO_FICHA.Fields("COD_FUNC_CENTRAL"), "00") & "'"
                            End If
                            
                            
                            If Format(W_ADO_FICHA.Fields("COD_FUNC_CENTRAL"), "00") = "99" Then MsgBox "Teste"
                            
                            w_COM = 0
                            w_Premio = 0
                            If Not de.rscmdSqlComissao.EOF Then w_COM = de.rscmdSqlComissao.Fields("TOTAL_COM")
                            If Not w_ado_Premio.EOF And w_ado_Premio.Filter <> 0 Then w_Premio = w_ado_Premio.Fields("Premio")
                            
                            'aa *** INCLUI PREMIO + COMISSÃO ***
                            If Not W_ADO_FICHA.EOF Then
                                
                                        W_DT_INI_MES = CVDate("01/" & TXT_MES & "/" & TXT_ANO)
                                        W_DT_FIM_MES = CVDate("01/" & Format(W_DT_INI_MES + 35, "MM/YYYY"))
                                        'sE DT DE ADM. FOR MAIOR Q/ A DT DO PRIMEIRO DIA DO MES ***
                                        If CVDate(W_ADO_FICHA.Fields("DT_ADM")) >= CVDate(W_DT_INI_MES) Then W_DT_INI_MES = CVDate(W_ADO_FICHA.Fields("DT_ADM"))
                                        
                                        '*** Se tem foi Demitido   então Pega QTde de dias trab.
                                        If Not IsNull(W_ADO_FICHA.Fields("DT_DEM")) Then
                                              W_QT_DIAS_TRAB = (CVDate(W_ADO_FICHA.Fields("DT_DEM")) + 1) - CVDate(W_DT_INI_MES)
                                        '*** Se a Adm é anterior ao mês da COmissão ***
                                        ElseIf W_DT_INI_MES = CVDate("01/" & TXT_MES & "/" & TXT_ANO) Then
                                              W_QT_DIAS_TRAB = "-30"
                                        '*** Se a foi Adm no meio do mês da COmissão ***
                                        Else
                                              W_QT_DIAS_TRAB = W_DT_FIM_MES - W_DT_INI_MES
                                        End If
                                        
                                        '*** Pega o Piso referente se for com ou sem registro
                                        w_Pdesc = IIf(IsNull(W_ADO_FICHA.Fields("DT_REG")), "Ps. B", "Ps. L")
                                        w_Piso = IIf(IsNull(W_ADO_FICHA.Fields("DT_REG")), W_ADO_FICHA.Fields("vpiso"), W_ADO_FICHA.Fields("vpiso_R"))
                                        If IsNull(w_Piso) Then
                                            w_Piso = 0
                                        End If
                                        W_VALOR_PISO = IIf(IsNull(W_ADO_FICHA.Fields("DT_REG")), W_ADO_FICHA.Fields("vpiso"), W_ADO_FICHA.Fields("vpiso_R"))
                                        w_Valor = 0
                                        
                                        '*** Pega os Valores Comissão e Premio e Total
                                        w_PRE = Format(0, "R$ 0.00")
                                        If Not w_ado_Premio.EOF And w_ado_Premio.Filter <> 0 Then w_PRE = Format(w_ado_Premio.Fields("Premio"), "R$ 0.00")
                                        w_COMI = Format(0, "R$ 0.00")
                                        If Not de.rscmdSqlComissao.EOF Then w_COMI = Format(de.rscmdSqlComissao.Fields("TOTAL_COM"), "R$ 0.00")
                                        W_TVnd = Format(0, "R$ 0.00")
                                        If Not de.rscmdSqlComissao.EOF Then W_TVnd = Format(de.rscmdSqlComissao.Fields("TOTAL_VND"), "R$ 0.00")
                                        w_TCP = Format(CDbl(w_PRE) + CDbl(w_COMI), "R$ 0.00")
                                                                        
                                
                                        '*** Se não Tiver conta 20,21,23 cadastrada ***
                                        If de.cnc.Execute("SELECT C_CODIGO FROM TAB_DESC_CALC WHERE C_N_FICHA = " & W_ADO_FICHA.Fields("M_NFICHA") & " AND C_TP_CONTA IN(20,21,23)").RecordCount = 0 Then
                                                If Not (Format(W_ADO_FICHA.Fields("COD_FUNC_CENTRAL"), "00") = "10" Or Format(W_ADO_FICHA.Fields("COD_FUNC_CENTRAL"), "00") = "99") Then
                                                    'b1*** Salva    TC  TP   e  ou  PISO ***
                                                    If W_QT_DIAS_TRAB = "-30" Then '*** Se for o Mês Completo ***
                                                        '*** Se o (TC + TP) > Piso bruto do mês inteiro *** então Insira TC  e TP
                                                        If (w_COM + w_Premio) > W_VALOR_PISO And W_ADO_FICHA.Fields("tipo") = "V" Then
                                                            If Not de.rscmdSqlComissao.EOF Then
                                                                de.cmdIncluirDescCalc Date, W_ADO_FICHA.Fields("M_NFICHA"), "20", "+", de.rscmdSqlComissao.Fields("TOTAL_COM"), "T. VND : " & Format(de.rscmdSqlComissao.Fields("TOTAL_VND"), "R$ 0.00"), 0, 0, 0, 0, 0
                                                                w_Valor = w_Valor + CDbl(de.rscmdSqlComissao.Fields("TOTAL_COM"))
                                                            End If
                                                            If Not w_ado_Premio.EOF Then
                                                                de.cmdIncluirDescCalc Date, W_ADO_FICHA.Fields("M_NFICHA"), "21", "+", CDbl(w_ado_Premio.Fields("Premio")), "T. Prêmio : " & Format(w_ado_Premio.Fields("Premio"), "R$ 0.00"), 0, 0, 0, 0, 0
                                                                w_Valor = w_Valor + CDbl(w_ado_Premio.Fields("Premio"))
                                                            End If
                                                        Else '*** Paga Piso Bruto do Mês ***
                                                            '*** Cria a Descriação ***
                                                            w_desc = "VND. - " & w_Pdesc & " : " & IIf(IsNull(W_VALOR_PISO), "R$ 0,00", Format(W_VALOR_PISO, "R$ 0.00")) & IIf(W_ADO_FICHA.Fields("COD_FUNC_CENTRAL") = "", " -  '' Código de Vendedor Não Cadastrado para Cálculo da Comissão! ''", "  ( T. Vnd.: " & W_TVnd & "  =  T.C (" & w_COMI & ") , T.P (" & w_PRE & ") = T. C+P (" & w_TCP & "))")
                                                            If IsNull(W_VALOR_PISO) Then W_VALOR_PISO = 0
                                                            de.cmdIncluirDescCalc Date, W_ADO_FICHA.Fields("M_NFICHA"), "23", "+", CDbl(W_VALOR_PISO), w_desc, 0, 0, 0, 0, 0
                                                            w_Valor = W_VALOR_PISO
                                                        End If
                                                        w_reg = w_reg + 1
                                                    ' Proporcional *** Se tiver alguns Dias trabalhados ***
                                                    ElseIf W_QT_DIAS_TRAB > 0 Then
                                                        w_Piso = IIf(IsNull(w_Piso), 0, w_Piso)
                                                        W_VALOR_PISO = W_QT_DIAS_TRAB * (CDbl(w_Piso) / 30)
                                                        
                                                        w_desc = "VND. - " & W_QT_DIAS_TRAB & " dias ref. ao " & w_Pdesc & " " & IIf(IsNull(w_Piso), "R$ 0,00", Format(w_Piso, "R$ 0.00")) & " :  < (" & Format(w_Piso, "R$ 0.00") & " / 30) = (" & Format(w_Piso / 30, "R$ 0.00") & " x " & W_QT_DIAS_TRAB & ") = " & Format(W_VALOR_PISO, "R$ 0.00") & ") >  " & IIf(W_ADO_FICHA.Fields("COD_FUNC_CENTRAL") = "", " -  '' Código de Vendedor Não Cadastrado para Cálculo da Comissão! ''", "  ( T. Vnd.: " & W_TVnd & "  =  T.C " & w_COMI & " )")
                                                            
                                                        '*** Se o TC > Piso Proporcional dos dias atrabalhados *** então Insira TC
                                                        If w_COM > W_VALOR_PISO And W_ADO_FICHA.Fields("tipo") = "V" Then
                                                            If Not de.rscmdSqlComissao.EOF Then de.cmdIncluirDescCalc Date, W_ADO_FICHA.Fields("M_NFICHA"), "20", "+", de.rscmdSqlComissao.Fields("TOTAL_COM"), w_desc, 0, 0, 0, 0, 0
                                                            w_Valor = de.rscmdSqlComissao.Fields("TOTAL_COM")
                                                        Else '*** Paga Piso Proporcional dos dias trabalhados ***
                                                            '*** Cria a Descriação ***
                                                            If IsNull(W_VALOR_PISO) Then W_VALOR_PISO = 0
                                                            de.cmdIncluirDescCalc Date, W_ADO_FICHA.Fields("M_NFICHA"), "23", "+", CDbl(W_VALOR_PISO), w_desc, 0, 0, 0, 0, 0
                                                            w_Valor = W_VALOR_PISO
                                                        End If
                                                        w_reg = w_reg + 1
                                                    
                                                    End If 'b1***
                                                Else
                                                            de.cmdIncluirDescCalc Date, W_ADO_FICHA.Fields("M_NFICHA"), "23", "=", 0, "", 0, 0, 0, 0, 0
                                                End If
                                        Else
                                            qt_ja_cad = qt_ja_cad + 1
                                        End If 'a1***
                                
                                If Not de.rscmdSqlComissao.EOF Then
                                    '*** Atualiza Valores de Referente a Comissões na ficha do Funcionario p/ ser somente usado no relatorio do PL
                                    '*** Atualiza o Saldo na Ficha Mensal do Funcionario
                                    If Not IsNumeric(W_TVnd) Then W_TVnd = 0
                                    de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_Vnd = '" & CDbl(W_TVnd) & "', M_TOTAL = M_Total + '" & CDbl(w_Valor) & "', M_TOTAL_MAIS = M_TOTAL_MAIS + '" & CDbl(w_Valor) & "', M_TOTAL_COM = '" & CDbl(w_COMI) & "', M_TOTAL_PRE = '" & CDbl(w_PRE) & "', M_TOTAL_VND = '" & W_TVnd & "'   WHERE (M_NFICHA = " & W_ADO_FICHA.Fields("M_NFICHA") & ")"
                                End If
                            End If 'aa***
                          'UPDATE M_COMISSÃO -
                          de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_COMISSAO = 'S' WHERE (M_NFICHA = " & W_ADO_FICHA.Fields("M_NFICHA") & ")"
                            
                    W_ADO_FICHA.MoveNext
                    Loop
                    


'***  entre os caixa **** calc a média ***
                        'filtra as fichas somente dos caixas
                        W_ADO_FICHA.Filter = "TIPO = 'C'"
                        If ck_Nome.value = 0 Then W_ADO_FICHA.Filter = "TIPO = 'C' AND M_F_COD = " & dbNome.BoundText & ""
                        
                        If Not W_ADO_FICHA.EOF Then W_ADO_FICHA.MoveFirst
                        Do While Not W_ADO_FICHA.EOF
                            If de.rscmdTotalVend.State = 1 Then de.rscmdTotalVend.Close
                            de.cmdTotalVend TXT_MES, TXT_ANO, W_ADO_FICHA.Fields("M_LOGO")
                            
                            '*** looping entre os Vendedores p/ Calc. Média
                            W_QT = 1
                            W_TT = 0
                            w_DESCR = ""
                            Do While Not de.rscmdTotalVend.EOF
                                W_TT = W_TT + de.rscmdTotalVend.Fields("valor")
                                w_DESCR = w_DESCR & IIf(w_DESCR = "", "< (" & Format(de.rscmdTotalVend.Fields("valor"), "0.00"), " + " & Format(de.rscmdTotalVend.Fields("valor"), "0.00"))
                                              
                                If W_QT = IIf(IsNull(W_ADO_FICHA.Fields("CX_QT_VND")), 3, W_ADO_FICHA.Fields("CX_QT_VND")) Then
                                    w_Media = W_TT / W_QT
                                    w_DESCR = w_DESCR & ") = " & Format(W_TT, "0.00") & " / " & W_QT & " = " & Format(w_Media, "0.00") & " >"
                                    Exit Do
                                End If
                                W_QT = W_QT + 1
                                de.rscmdTotalVend.MoveNext
                            Loop
                            
                            
                                                
                            '*** Pega o Piso referente se for com ou sem registro
                            If IsNull(W_ADO_FICHA.Fields("Dt_Reg")) Then
                                w_Piso = W_ADO_FICHA.Fields("vpiso")
                                w_Pdesc = "Ps. B"
                            Else
                                w_Piso = W_ADO_FICHA.Fields("vpiso_R")
                                w_Pdesc = "Ps. L"
                            End If
                            w_Piso = IIf(IsNull(w_Piso), 0, w_Piso)
                            
                            '*** paga comissão *** da média
                            If w_Media > w_Piso Then
                                 w_desc = "CX - " & w_Pdesc & " : " & IIf(IsNull(w_Piso), "R$ 0,00", Format(w_Piso, "R$ 0.00")) & "   " & w_DESCR
                                de.cmdIncluirDescCalc Date, W_ADO_FICHA.Fields("M_NFICHA"), "22", "+", w_Media, w_desc, 0, 0, 0, 0, 0
                                W_REG_CX = W_REG_CX + 1
                            '*** paga piso ***
                            Else
                            
                                        W_DT_INI_MES = CVDate("01/" & TXT_MES & "/" & TXT_ANO)
                                        W_DT_FIM_MES = CVDate("01/" & Format(W_DT_INI_MES + 35, "MM/YYYY"))
                                        'sE DT DE ADM. FOR MAIOR Q/ A DT DO PRIMEIRO DIA DO MES ***
                                        If CVDate(W_ADO_FICHA.Fields("DT_ADM")) >= CVDate(W_DT_INI_MES) Then
                                             W_DT_INI_MES = CVDate(W_ADO_FICHA.Fields("DT_ADM"))
                                        End If
                                        
                                        If Not IsNull(W_ADO_FICHA.Fields("DT_DEM")) Then
                                              W_QT_DIAS_TRAB = (CVDate(W_ADO_FICHA.Fields("DT_DEM")) + 1) - CVDate(W_DT_INI_MES)
                                        ElseIf W_DT_INI_MES = CVDate("01/" & TXT_MES & "/" & TXT_ANO) Then
                                              W_QT_DIAS_TRAB = "-30"
                                        Else
                                              W_QT_DIAS_TRAB = W_DT_FIM_MES - W_DT_INI_MES
                                              
                                        End If
                                        
                                        
                                        '*** INCLUI PISO S/ REGISTRO ***
                                        If IsNull(W_ADO_FICHA.Fields("DT_REG")) Then
                                            If W_QT_DIAS_TRAB = "-30" Then
                                                W_VALOR_PISO = W_ADO_FICHA.Fields("vpiso")
                                                w_desc = "CX - " & w_Pdesc & " : " & IIf(IsNull(W_VALOR_PISO), "R$ 0,00", Format(W_VALOR_PISO, "R$ 0.00")) & "   " & w_DESCR
                                            Else
                                                W_VALOR_PISO = W_QT_DIAS_TRAB * (W_ADO_FICHA.Fields("vpiso") / 30)
                                                w_desc = "CX - " & W_QT_DIAS_TRAB & " dias ref. ao " & w_Pdesc & " " & IIf(IsNull(W_VALOR_PISO), "R$ 0,00", Format(W_ADO_FICHA.Fields("vpiso"), "R$ 0.00")) & " :  (" & Format(W_ADO_FICHA.Fields("vpiso"), "R$ 0.00") & " / 30 = " & Format(W_ADO_FICHA.Fields("vpiso") / 30, "R$ 0.00") & " x " & W_QT_DIAS_TRAB & ")"
                                            End If
                                                
                                            de.cmdIncluirDescCalc Date, W_ADO_FICHA.Fields("M_NFICHA"), "22", "+", CDbl(W_VALOR_PISO), w_desc, 0, 0, 0, 0, 0
                                            W_REG_CX = W_REG_CX + 1
                                            
                                        '*** INCLUI PISO C/ REGISTRO ***
                                        Else
                                            If W_QT_DIAS_TRAB = "-30" Then
                                                W_VALOR_PISO = W_ADO_FICHA.Fields("vpiso_R")
                                                w_desc = "CX - " & w_Pdesc & " : " & IIf(IsNull(W_VALOR_PISO), "R$ 0,00", Format(W_VALOR_PISO, "R$ 0.00")) & "   " & w_DESCR
                                            Else
                                                W_VALOR_PISO = W_QT_DIAS_TRAB * (W_ADO_FICHA.Fields("vpiso_R") / 30)
                                                w_desc = "CX - " & W_QT_DIAS_TRAB & " dias ref. ao " & w_Pdesc & " " & IIf(IsNull(W_VALOR_PISO), "R$ 0,00", Format(W_ADO_FICHA.Fields("vpiso_R"), "R$ 0.00")) & " :  (" & Format(W_ADO_FICHA.Fields("vpiso_R"), "R$ 0.00") & " / 30) = " & Format(W_ADO_FICHA.Fields("vpiso_R") / 30, "R$ 0.00") & " x " & W_QT_DIAS_TRAB & ")"
                                            End If
                                            
                                            If IsNull(W_VALOR_PISO) Then W_VALOR_PISO = 0
                                            
                                            de.cmdIncluirDescCalc Date, W_ADO_FICHA.Fields("M_NFICHA"), "22", "+", CDbl(W_VALOR_PISO), w_desc, 0, 0, 0, 0, 0
                                            
                                            W_REG_CX = W_REG_CX + 1
                                        End If
                                    End If
                            
                                                        
                            W_ADO_FICHA.MoveNext
                        Loop
                    
                    
                    
                    MsgBox w_reg & " comissões de funcionários criadas com sucesso!" & Chr(13) & W_REG_CX & " S. de Cx. criados com sucesso!", vbInformation
                    w_reg = 0
                    
                    
            Else
                MsgBox "O SISTEMA NÃO PODE GERAR COMISSÃO SEM Q/ ANTES SEJA GERADO AS FICHAS NESTE PERÍODO!", vbCritical
            End If
        Else
            MsgBox "Preencha os Campos!", vbCritical
        End If
    
    
ElseIf V_RESP = vbYes Then
    MsgBox "Já foram criadas as comissão nesta data!", vbCritical
End If
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
    
End Sub









Private Sub ck_Logo_Click()
    If ck_Logo.value = 1 Then
        TXT_LOGO = "%"
    End If
End Sub

Private Sub ck_Nome_Click()
    If ck_Nome.value = 1 Then
        dbNome = "%"
        dbNome.Enabled = False
    Else
        dbNome = ""
        dbNome.Enabled = True
        On Error Resume Next
        dbNome.SetFocus
        SendKeys "{f4}"
    End If
End Sub

Private Sub Form_Load()
    TXT_MES = CDbl(Format(Date, "mm"))
    TXT_ANO = Format(Date, "yyyy")
    
    Set ADO_FUNC.Recordset = de.cnc.Execute("SELECT * FROM TAB_FUNCIONARIO WHERE NOT F_NOME = '10 - Func' and not F_NOME = '99 - Presence' ORDER BY F_NOME").Clone
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Fechar
End Sub

Private Sub TXT_ANO_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub TXT_ANO_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub TXT_LOGO_Change()
On Error Resume Next
    
    Set ADO_FUNC.Recordset = de.cnc.Execute("SELECT * FROM TAB_FUNCIONARIO WHERE NOT F_NOME = '10 - Func' and not F_NOME = '99 - Presence' AND F_COD_L = '" & TXT_LOGO & "' ORDER BY F_NOME ").Clone
    dbNome.ReFill
    dbNome.Refresh
    

End Sub

Private Sub TXT_LOGO_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub TXT_MES_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"
End Sub

'--------- Ao Pressionar uma Tecla -----------
Private Sub TXT_mes_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_ano_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub TXT_MES_Validate(Cancel As Boolean)
    If Not (CDbl(IIf(Not (IsNumeric(TXT_MES)), 0, TXT_MES)) >= 1 And CDbl(IIf(Not (IsNumeric(TXT_MES)), 0, TXT_MES)) <= 12) Then
        MsgBox "Você deve digitar o Nº Mês!", vbInformation
        TXT_MES.SetFocus
    End If
End Sub


' -------  Teclas de Atalhos --------
Sub Keys(KeyCode As Integer, Shift As Integer)
    '*** Shift (4 = Alt) ***
    If Shift = 4 Then
        Select Case KeyCode
        Case 70: ' "F"
                Fechar
        Case 71: ' "G"
                Gerar
        End Select
    End If
End Sub

