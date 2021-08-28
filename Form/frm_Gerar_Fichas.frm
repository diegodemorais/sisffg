VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "msCOMCTL.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form frm_Gerar_Fichas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerar Fichas"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   Icon            =   "frm_Gerar_Fichas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   3135
   StartUpPosition =   2  'CenterScreen
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
      ItemData        =   "frm_Gerar_Fichas.frx":08CA
      Left            =   240
      List            =   "frm_Gerar_Fichas.frx":08F2
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
      Left            =   1005
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
      Width           =   3135
      _ExtentX        =   5530
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
               Picture         =   "frm_Gerar_Fichas.frx":091D
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Gerar_Fichas.frx":0C37
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Gerar_Fichas.frx":0F51
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Gerar_Fichas.frx":126B
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSDataListLib.DataCombo TXT_LOGO 
      Bindings        =   "frm_Gerar_Fichas.frx":1B45
      DataField       =   "F_COD_L"
      DataSource      =   "ADOREG"
      Height          =   405
      Left            =   2160
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
      Left            =   2160
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
      Left            =   360
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
      Left            =   1200
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   120
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "frm_Gerar_Fichas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BarraF_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err1
   
    Select Case Button.key
        Case "fechar": Fechar
        Case "gerar": Gerar
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
    If de.rsTAB_FICHA_MENS.State = 1 Then de.rsTAB_FICHA_MENS.Requery
    Unload Me
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Gerar()
On Error GoTo err1

'*** CONFIRMA GERAÇÃO DE MESES DIFERENTES AO ATUAL ***
V_RESP = vbYes
If TXT_ANO <> Format(Date, "YYYY") Or TXT_MES <> CDbl(Format(Date, "MM")) Then V_RESP = MsgBox("DESEJA REALMENTE GERAR UM MÊS QUE NÃO SEJA O ATUAL?", vbQuestion + vbYesNo + vbDefaultButton2)


If de.rsTAB_FICHA_MENS.State = 0 Then de.TAB_FICHA_MENS

'*** Verifica se existe algo cadastrado ***
If V_RESP = vbYes And de.cnc.Execute("SELECT COUNT(M_ANO) AS Qtde FROM TAB_FICHA_MENS WHERE (M_MES = " & CDbl(TXT_MES) & ") AND (M_ANO = " & CDbl(TXT_ANO) & ") and M_Logo Like '" & TXT_LOGO & "%'").Fields("Qtde") = 0 Then
    
    If TXT_ANO <> "" And TXT_MES <> "" Then
        
        '*** BLOQUEIA AS FICHAS PASSADAS ****
        w_mes = CDbl(TXT_MES) - 1
        If w_mes <= 0 Then
            w_mes = 12 + w_mes
            w_ano = TXT_ANO - 1
        End If
        
        'Bloqueia as fichas de  2 meses p/ traz
        'de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_BLOQ = " & -1 & " Where (M_Logo Like '" & TXT_LOGO & "%' and m_mes = " & w_mes & ")"
        
        'Atualização dos Totais das fichas
        If de.rsTAB_FICHA_MENS.State = 1 Then de.rsTAB_FICHA_MENS.Close
        de.TAB_FICHA_MENS
        'FILTRA AS FICHAS DO MÊS E Q/ NÃO ESTA BLOQUEADA
        de.rsTAB_FICHA_MENS.Filter = "M_MES = " & IIf(w_mes = 12, 1, w_mes + 1) & " AND M_BLOQ = 0"
        
        On Error Resume Next
        Do While Not de.rsTAB_FICHA_MENS.EOF
            W_MAIS = 0
            W_MENOS = 0
            W_TOTAL = 0
            '*** CALCULA O TOTAL - APÓS O NOVO VALOR ***
            W_MAIS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MAIS FROM TAB_DESC_CALC  WHERE (C_TP_OP = '+') AND (C_N_FICHA = " & de.rsTAB_FICHA_MENS.Fields("M_NFICHA") & ")").Fields("MAIS")
            W_MENOS = de.cnc.Execute("SELECT SUM(C_VALOR) AS MENOS FROM TAB_DESC_CALC WHERE (C_TP_OP = '-') AND (C_N_FICHA = " & de.rsTAB_FICHA_MENS.Fields("M_NFICHA") & ")").Fields("MENOS")
            W_TOTAL = IIf(IsNull(W_MENOS), 0, W_MENOS) + IIf(IsNull(W_MAIS), 0, W_MAIS)
            '*** Atualiza os Campos  Total , Mais e Menos
            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_TOTAL = '" & CDbl(IIf(IsNull(W_TOTAL), 0, W_TOTAL)) & "', M_TOTAL_MAIS = '" & CDbl(IIf(IsNull(W_MAIS), 0, W_MAIS)) & "', M_TOTAL_MENOS = '" & CDbl(IIf(IsNull(W_MENOS), 0, W_MENOS)) & "' WHERE (M_NFICHA = " & de.rsTAB_FICHA_MENS.Fields("M_NFICHA") & ")"
        
        de.rsTAB_FICHA_MENS.MoveNext
        Loop
                
On Error GoTo err1
    Dim w_mes_ger, w_ano_ger As Integer
         
         w_mes_ger = IIf((TXT_MES - 1) = 0, 12, TXT_MES - 1)
         w_ano_ger = IIf((TXT_MES - 1) = 0, TXT_ANO - 1, TXT_ANO)
        
        
        'Gera as fichas desejadas
        'de.cnc.Execute "INSERT INTO `TAB_FICHA_MENS` (M_F_COD, M_ANO, M_MES, M_Ferias, M_OBS, M_NOTAS, M_Logo, M_ANOTACAO, M_DT_ADM, M_DT_REG, M_DT_DEM, M_Nome, M_FERIAS_PG, M_FERIAS_ULT_PG, M_FERIAS_OK, M_13_PG, M_13_ULT_PG, M_13_OBS, M_13_OK, M_DEM_OK, M_VCTO_FERIAS, M_PG_SAL_FAM, M_NUM_FILHOS, M_PG_VT, M_TIPO) " & _
        '                                    "SELECT F_Codigo, " & TXT_ANO & " AS ano, " & TXT_MES & " AS mes, F_Ferias, F_OBS, F_NOTAS, F_Cod_L, F_ANOTACAO, F_DT_ADM, F_DT_REG, F_DT_DEM, F_NOME, F_FERIAS_PG, F_FERIAS_ULT_PG, F_FERIAS_OK, F_13_PG, F_13_ULT_PG, F_13_OBS, F_13_OK, F_DEM_OK, F_VCTO_FERIAS, F_PG_SAL_FAM, F_NUM_FILHOS, F_PG_VT, F_TIPO FROM TAB_FUNCIONARIO WHERE (F_Cod_L Like '" & TXT_LOGO & "%' and F_DT_DEM IS NULL)", w_reg
        
        de.cnc.Execute "INSERT INTO `TAB_FICHA_MENS` (M_F_COD, M_ANO, M_MES, M_Ferias, M_OBS, M_NOTAS, M_Logo, M_ANOTACAO, M_DT_ADM, M_DT_REG, M_DT_DEM, M_Nome, M_FERIAS_PG, M_FERIAS_ULT_PG, M_FERIAS_OK, M_13_PG, M_13_ULT_PG, M_13_OBS, M_13_OK, M_DEM_OK, M_VCTO_FERIAS, M_PG_SAL_FAM, M_NUM_FILHOS, M_PG_VT, M_TIPO) " & _
                                            "SELECT F_Codigo, " & TXT_ANO & " AS ano, " & TXT_MES & " AS mes, F_Ferias, F_OBS, F_NOTAS, F_Cod_L, F_ANOTACAO, F_DT_ADM, F_DT_REG, F_DT_DEM, F_NOME, F_FERIAS_PG, F_FERIAS_ULT_PG, F_FERIAS_OK, F_13_PG, F_13_ULT_PG, F_13_OBS, F_13_OK, F_DEM_OK, F_VCTO_FERIAS, F_PG_SAL_FAM, F_NUM_FILHOS, F_PG_VT, F_TIPO FROM TAB_FUNCIONARIO INNER JOIN TAB_FICHA_MENS ON TAB_FUNCIONARIO.F_Codigo = TAB_FICHA_MENS.M_F_COD WHERE (TAB_FUNCIONARIO.F_Cod_L Like '" & TXT_LOGO & "%' and TAB_FICHA_MENS.M_DT_DEM IS NULL) AND (TAB_FICHA_MENS.M_ANO= " & w_ano_ger & " AND TAB_FICHA_MENS.M_MES=" & w_mes_ger & ")", w_reg
        
        
        de.rsTAB_FICHA_MENS.Requery
        Pause 0.5
    
    '*** Gerar Lançamentos Fixos
    
        Dim adoFuncs As ADODB.Recordset
        Dim adoFixos As ADODB.Recordset
        
        Set adoFuncs = de.cnc.Execute("SELECT DISTINCT(TAB_FICHA_MENS.M_NFICHA) as FICHA, TAB_FICHA_MENS.M_F_COD as COD FROM TAB_FICHA_MENS, TAB_DESC_CALC_FIXO WHERE TAB_FICHA_MENS.M_F_COD = TAB_DESC_CALC_FIXO.CF_EMP_COD AND (TAB_FICHA_MENS.M_ANO = " & TXT_ANO & ") AND (TAB_FICHA_MENS.M_MES = " & TXT_MES & ")").Clone
    
        Do While Not adoFuncs.EOF
            Set adoFixos = de.cnc.Execute("SELECT * FROM TAB_DESC_CALC_FIXO WHERE CF_EMP_COD = " & adoFuncs.Fields("COD")).Clone
        
            Do While Not adoFixos.EOF
                de.cmdIncluirDescCalc2 adoFixos.Fields("CF_DT"), adoFuncs.Fields("FICHA"), adoFixos.Fields("CF_TP_CONTA"), adoFixos.Fields("CF_TP_OP"), adoFixos.Fields("CF_VALOR"), adoFixos.Fields("CF_DESC"), "0", adoFixos.Fields("CF_CODIGO"), "0", "0", adoFixos.Fields("CF_EMP_COD"), 0
                adoFixos.MoveNext
            Loop
          
            adoFuncs.MoveNext
        Loop
    
                
    '***  Gerar Salario Familia
        'Pega os registro das  fichas  geradas , cujo o mes  e  ano e loja sejam  iguais  e  que  sal_familia  esteje  ligado
        Dim adoFichas As ADODB.Recordset
        Set adoFichas = de.cnc.Execute("SELECT M_NFICHA, M_PG_SAL_FAM, M_NUM_FILHOS, M_LOGO   FROM TAB_FICHA_MENS WHERE (M_LOGO Like '" & TXT_LOGO & "%' and m_Mes = " & TXT_MES & " and m_ano = " & TXT_ANO & " and M_PG_SAL_FAM = -1)").Clone
        
        wSalFam = de.cnc.Execute("Select Sal_Familia from tab_config").Fields(0)
        
        On Error Resume Next
        Do While Not adoFichas.EOF
            wValor = 0
            wValor = Format(adoFichas.Fields("m_num_filhos") * wSalFam, "0.00") 'Calcula Salario
            wDesc = "(" & Format(wSalFam, "0.00") & " x " & adoFichas.Fields("m_num_filhos") & ") = " & Format(wValor, "0.00")
            de.cmdIncluirDescCalc Date, adoFichas.Fields("M_NFicha"), 26, "+", wValor, wDesc, "", "0", "0", "0", "0"
        
            adoFichas.MoveNext
        Loop
                
                
    '**************************** Sal.  Fam.
    
                
'*****  PRESTAÇÕES DE EMRPESTIMO ****
            If de.rsTAB_FICHA_MENS.State = 1 Then de.rsTAB_FICHA_MENS.Close
            de.TAB_FICHA_MENS
            
            Dim W_ADO_EMP As ADODB.Recordset
            'Zera a descrição dos q/ tem saldo zero
            Set W_ADO_EMP = de.cnc.Execute("SELECT * FROM TAB_EMPRESTIMO, TAB_Funcionario WHERE E_SALDO = 0 and TAB_EMPRESTIMO.E_F_COD = TAB_Funcionario.F_CODIGO AND TAB_Funcionario.F_COD_L LIKE '" & TXT_LOGO & "%'").Clone
            Do While Not W_ADO_EMP.EOF
                '*** Dá baixa no emprestimo na tab. funcionario ***
                de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_EMPRESTIMO_ANOT = '' WHERE (F_Codigo = " & W_ADO_EMP.Fields("E_F_COD") & ")"
                W_ADO_EMP.MoveNext
            Loop
            
            Set W_ADO_EMP = de.cnc.Execute("SELECT * FROM TAB_EMPRESTIMO, TAB_Funcionario WHERE E_SALDO > 0 and TAB_EMPRESTIMO.E_F_COD = TAB_Funcionario.F_CODIGO AND TAB_Funcionario.F_COD_L LIKE '" & TXT_LOGO & "%'").Clone
            
            Do While Not W_ADO_EMP.EOF
                '*** CALCULA SOMENTE SE EXISTIR FICHA NESTE MÊS -****
                If de.cnc.Execute("SELECT M_NFICHA FROM TAB_FICHA_MENS WHERE M_F_COD = " & W_ADO_EMP.Fields("E_F_COD") & " AND M_MES = " & TXT_MES & " AND M_ANO = " & TXT_ANO & "").RecordCount > 0 Then
                        w_parc = de.cnc.Execute("Select EP_Parc from tab_Emprestimo_pg Where ep_codigo = " & W_ADO_EMP.Fields("E_codigo") & " and ep_parc > 0 ").RecordCount + 1
                        W_DT_PG = CVDate("01/" & TXT_MES & "/" & TXT_ANO) + 32
                        
                        If IsDate((W_ADO_EMP.Fields("E_DIA_PG") & "/" & TXT_MES & "/" & TXT_ANO)) Then
                            W_DT_PG = CVDate(W_ADO_EMP.Fields("E_DIA_PG") & "/" & TXT_MES & "/" & TXT_ANO) + 31
                            W_DT_PG = CVDate(W_ADO_EMP.Fields("E_DIA_PG") & "/" & Format(W_DT_PG, "mm/yyyy"))
                        Else
                            W_DT_PG = CVDate("01/" & TXT_MES & "/" & TXT_ANO) + 32
                            W_DT_PG = CVDate("01/" & Format(W_DT_PG, "mm/yyyy")) - 1
                            If CDbl(Format(W_DT_PG, "dd")) < W_ADO_EMP.Fields("E_DIA_PG") Then
                                w_QtDias = W_ADO_EMP.Fields("E_DIA_PG") - CDbl(Format(W_DT_PG, "dd"))
                            End If
                            W_DT_PG = W_DT_PG + w_QtDias
                        End If
                        
                        
                        W_JUROS = Format(CALC_PG_EMP(W_ADO_EMP, W_DT_PG), "R$ 0.00")
                        w_Valor = (W_ADO_EMP.Fields("E_SALDO") / IIf(W_PARC_RESTANTE = 0, 1, W_PARC_RESTANTE)) + CDbl(W_JUROS)
                                                
                        W_NFICHA = de.cnc.Execute("SELECT M_NFICHA FROM TAB_FICHA_MENS WHERE M_F_COD = " & W_ADO_EMP.Fields("E_F_COD") & " AND M_MES = " & TXT_MES & " AND M_ANO = " & TXT_ANO & "").Fields(0)
                        
                        W_DESC_CONTA = "Pg. Emp.: " & W_ADO_EMP.Fields("E_QT_PG") + 1 & "/" & W_ADO_EMP.Fields("E_QT_PARC") & vbCrLf & "Valor : " & Format(w_Valor - W_JUROS, "R$ 0.00") & "    +    Juros : " & Format(W_JUROS, "R$ 0.00")
                        
                        '*** INCLUI CONTA P/ DESCONTO DO EMP. ***
                        de.cmdIncluirDescCalc W_DT_PG, W_NFICHA, "9", "-", CDbl(w_Valor * -1), W_DESC_CONTA, "0", "0", CDbl(W_JUROS), w_parc, W_ADO_EMP.Fields("E_CODIGO")
                        '*** iNCLUINDO PAGAMENTO DE EMPRESTIMO  -  TAB_EMPRESTIMO_PG ***
                        W_C_COD = de.cnc.Execute("SELECT MAX(C_CODIGO)AS COD FROM TAB_DESC_CALC WHERE C_N_FICHA = " & W_NFICHA & "").Fields(0)
                        de.cmdIncluirEmprestimoPG W_ADO_EMP.Fields("E_CODIGO"), W_DT_PG, w_parc, w_qt_dias, CDbl(CDbl(w_Valor) - CDbl(W_JUROS)), CDbl(W_JUROS), W_C_COD
            
                        '*** Dá baixa no emprestimo na tab. funcionario ***
                        de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_EMPRESTIMO = F_EMPRESTIMO - '" & CDbl(w_Valor - W_JUROS) & "' WHERE (F_Codigo = " & W_ADO_EMP.Fields("E_F_COD") & ")"
                        
                        '*** Dá baixa no emprestimo na tab. emprestimo ***
                        de.cnc.Execute "UPDATE TAB_EMPRESTIMO SET E_QT_PG = E_QT_PG + 1 , E_DT_ULT_PG = '" & W_DT_PG & "', E_SALDO = E_SALDO - '" & CDbl(w_Valor - W_JUROS) & "' WHERE (E_Codigo = " & W_ADO_EMP.Fields("E_CODIGO") & ")"
                        
                        '*** ATUALIZAR A ANOTAÇÃO DO EMPRESTIMO DO FUNCIONARIO ***
                        '** Sql EMP. P/ GRID
                            
                            W_EMP_ANOT = ""
                            Dim ADO_ANOT As ADODB.Recordset
                            
                            w_Dt = CVDate("01/" & TXT_MES & "/" & TXT_ANO) + 65
                            Set ADO_ANOT = de.cnc.Execute("SELECT * FROM TAB_EMPRESTIMO WHERE E_F_COD = " & W_ADO_EMP.Fields("E_F_COD") & " AND (E_SALDO > 0  OR E_DT_ULT_PG <= #" & Format(w_Dt, "MM/DD/YYYY") & "#)").Clone
                            Do While Not ADO_ANOT.EOF
                                W_EMP_ANOT = W_EMP_ANOT & IIf(Len(W_EMP_ANOT) > 0, vbCrLf, "") & ". Dt Emp.: " & ADO_ANOT.Fields("E_DT_EMP") & "    Valor Emp.: " & Format(ADO_ANOT.Fields("E_VALOR"), "R$ 0.00") & "     Juros : " & ADO_ANOT.Fields("E_Juro_ao_mes") * 100 & " %" & "     Parc. Pg.: " & ADO_ANOT.Fields("E_QT_PG") & " / " & ADO_ANOT.Fields("E_QT_PARC")
                                W_EMP_ANOT = W_EMP_ANOT & vbCrLf & ". Saldo Ant.: " & Format(W_ADO_EMP.Fields("E_SALDO"), "R$ 0.00") & "         Dt Ult. Pg.: " & ADO_ANOT.Fields("E_DT_ULT_PG") & "        Saldo At.: " & Format(ADO_ANOT.Fields("E_SALDO"), "R$ 0.00")
                            
                                ADO_ANOT.MoveNext
                            Loop
                            
                            
                            '*** UPDATE NO FUNCIONARIO ATUALIZANDO A ANOTAÇÃO DO EMPRESTIMO ***
                            de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_EMPRESTIMO_ANOT = '" & W_EMP_ANOT & "' WHERE (F_Codigo = " & W_ADO_EMP.Fields("E_F_COD") & ")"
                
                            '*** Atualiza o Valor Total da Ficha ***
                            de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_TOTAL = M_TOTAL + '" & w_Valor & "', M_EMPRESTIMO_ANOT = '" & IIf(W_EMP_ANOT = "", " ", W_EMP_ANOT) & "' WHERE (M_NFICHA = " & W_NFICHA & ")"
                End If
            
                W_ADO_EMP.MoveNext
            Loop
            
            Set W_ADO_EMP = Nothing
               
         'PEGA O VALOR DA FICHA DO MÊS ANTERIOR  E  SE O SALDO FOR NEGATIVO    JOGA NO CAD. FUNC. COMO SALDO DEVEDOR ANTERIOR ***
         'ATUALIZA OS VALORES DEVEDORES DA FICHA NO CAD. DE FUNCIONARIO P/ Q/ POSSA SER MANTIDO O VALOR P/ PROXIMA FICHA
         w_mm = IIf((TXT_MES - 1) = 0, 12, TXT_MES - 1)
         w_yy = IIf((TXT_MES - 1) = 0, TXT_ANO - 1, TXT_ANO)
         
         'Atualizas na cad. Func
        de.cnc.Execute "UPDATE TAB_FUNCIONARIO INNER JOIN TAB_FICHA_MENS ON TAB_FUNCIONARIO.F_Codigo = TAB_FICHA_MENS.M_F_COD SET TAB_FUNCIONARIO.F_SALDO_ANT = [TAB_FICHA_MENS].[M_TOTAL] WHERE (((TAB_FICHA_MENS.M_MES)=" & w_mm & ") AND ((TAB_FICHA_MENS.M_ANO)=" & w_yy & ") AND ((TAB_FICHA_MENS.M_TOTAL)<0));"

                
'*****  PRESTAÇÕES DO SALDO DEVEDOR ANTERIOR, DA FICHA DO MÊS PASSADO ****
'            Dim W_ADO_SALDO As ADODB.Recordset
'            If de.rsTAB_FUNCIONARIO.State = 1 Then de.rsTAB_FUNCIONARIO.Requery
'
'            Set W_ADO_SALDO = de.cnc.Execute("SELECT * FROM TAB_FUNCIONARIO WHERE F_SALDO_ANT < 0 AND F_COD_L LIKE '" & TXT_LOGO & "%'").Clone
'
'            Do While Not W_ADO_SALDO.EOF
'
'            If W_ADO_SALDO.Fields("F_COD_L") <> "RP" Then  'Se não for Fichas RP
'
'
'                '*** CALCULA SOMENTE SE EXISTIR FICHA NESTE MÊS -****
'                If de.cnc.Execute("SELECT M_NFICHA FROM TAB_FICHA_MENS WHERE M_F_COD = " & W_ADO_SALDO.Fields("F_CODIGO") & " AND M_MES = " & TXT_MES & " AND M_ANO = " & TXT_ANO & "").RecordCount > 0 Then
'
'                        '*** PEGA O Nº DA FICHA GERADA NOVA ***
'                        W_NFICHA = de.cnc.Execute("SELECT M_NFICHA FROM TAB_FICHA_MENS WHERE M_F_COD = " & W_ADO_SALDO.Fields("F_CODIGO") & " AND M_MES = " & TXT_MES & " AND M_ANO = " & TXT_ANO & "").Fields(0)
'
'                        '*** SALDO DEVEDOR ***
'                        w_Valor = CDbl(W_ADO_SALDO.Fields("F_SALDO_ANT"))
'
'                        '*** DESCRIÇÃO DA CONTA ***
'                        W_DESC_CONTA = "Pg. Saldo Dev.: " & Format(w_Valor, "R$ 0.00")
'
'                        '*** INCLUI CONTA P/ DESCONTO DO EMP. *** e '*** VISTA CONTA CRIADA P/ NÃO PODER DELETAR ***
'                        de.cmdIncluirDescCalcVistado "01/" & TXT_MES & "/" & TXT_ANO, W_NFICHA, "14", "-", w_Valor, W_DESC_CONTA, 0, 0, 0, 0, 0
'
'                        '*** Dá baixa no SALDO DEVEDOR  na tab. funcionario ***
'                        de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_SALDO_ANT = 0 WHERE (F_Codigo = " & W_ADO_SALDO.Fields("F_CODIGO") & ")"
'
'                        '*** Atualiza o Valor Total da Ficha ***
'                        de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_TOTAL = M_TOTAL + '" & w_Valor & "' WHERE (M_NFICHA = " & W_NFICHA & ")"
'
'                End If
'
'                W_ADO_SALDO.MoveNext
'
'
'                End If 'Fichas RP
'
'            Loop
'
'            Set W_ADO_SALDO = Nothing

'SALDO NEGATIVO DO MÊS PASSADO
    'Dim vrVenda, vrFixo, vrMinimo, percComis, vrSalario, vrComis
    'Dim ww_mes, ww_ano, qtdeSaldoAdded
    'Dim adoFichasSaldo As ADODB.Recordset
            
    'ww_mes = TXT_MES
    'ww_ano = TXT_ANO
    
    'TXT_MES = adoFichasSaldo.Fields("M_MES") - 1
    'If TXT_MES = 0 Then
    '    TXT_MES = 12
    '    TXT_ANO = adoFichasSaldo.Fields("M_ANO") - 1
    'Else
    '    TXT_ANO = adoFichasSaldo.Fields("M_ANO")
    'End If
            
    'Set adoFichasSaldo = de.cnc.Execute("SELECT * FROM TAB_FICHA_MENS WHERE (M_ANO = " & TXT_ANO & ") AND (M_MES = " & TXT_MES & ")").Clone
        
  'On Error Resume Next

    'qtdeSaldoAdded = 0
    
    'Dim ADO_TOTAL As ADODB.Recordset
    'Dim wTXT_MAIS
    'Dim wTXT_MENOS
    'Dim wTXT_TOTAL
   
    'adoFichasSaldo.MoveFirst
    'Do While Not adoFichasSaldo.EOF

        
     ' wTXT_MAIS = 0
     ' wTXT_MENOS = 0
     ' wTXT_TOTAL = 0
      
      'Set ADO_TOTAL = de.cnc.Execute("SELECT tab_desc_calc.C_DT AS DATA, TAB_DESC_CALC.C_TP_CONTA AS TP, TAB_TP_CONTA.TP_DESC+' :: '+TAB_DESC_CALC.C_DESC AS CONTA, tab_desc_calc.C_VALOR AS VALOR, TAB_DESC_CALC.C_TP_OP AS OP, tab_desc_calc.C_VISTO AS VISTO, TAB_DESC_CALC.C_CODIGO, TAB_DESC_CALC.C_NCRED " _
      '                & "From TAB_DESC_CALC, TAB_TP_CONTA Where (TAB_DESC_CALC.C_TP_CONTA = [TAB_TP_CONTA].[TP_COD] And (TAB_DESC_CALC.C_N_FICHA = " & adoFichasSaldo.Fields("M_NFICHA") & ")) ORDER BY TAB_DESC_CALC.C_TP_OP, C_DT").Clone
                      
      
     ' If Not ADO_TOTAL.EOF Then
     '     ADO_TOTAL.MoveFirst
     '     Do While Not ADO_TOTAL.EOF
     '         If ADO_TOTAL.Fields("VALOR") >= 0 And ADO_TOTAL.Fields("OP") = "+" Then
     '             wTXT_MAIS = CDbl(wTXT_MAIS) + ADO_TOTAL.Fields("VALOR")
     '         ElseIf ADO_TOTAL.Fields("VALOR") < 0 And ADO_TOTAL.Fields("OP") = "-" Then
     '             wTXT_MENOS = CDbl(wTXT_MENOS) + ADO_TOTAL.Fields("VALOR")
     '         End If
     '         ADO_TOTAL.MoveNext
     '     Loop
          
     '     wTXT_TOTAL = CDbl(wTXT_MAIS) + CDbl(wTXT_MENOS)
     ' End If
    
     ' Dim proxFicha
     ' Dim w_desc
      
     ' proxFicha = de.cnc.Execute("SELECT M_NFICHA From TAB_FICHA_MENS WHERE M_ANO = " & ww_ano & " AND M_MES = " & ww_mes & " AND M_F_COD = " & adoFichasSaldo.Fields("M_F_COD")).Fields(0)
     ' de.cnc.Execute ("DELETE FROM TAB_DESC_CALC Where (C_TP_CONTA = 14) And (C_N_FICHA = " & proxFicha & ")")
      'If wTXT_TOTAL < 0 And adoFichasSaldo.Fields("M_LOGO") <> "RP" And (IsNull(adoFichasSaldo.Fields("M_DT_ACF")) Or adoFichasSaldo.Fields("M_DT_ACF") = "") Then
      'If wTXT_TOTAL < 0 And ADOREG.Recordset.Fields("M_LOGO") <> "RP" And (IsNumeric(proxFicha)) Then
      'If wTXT_TOTAL < 0 And ADOREG.Recordset.Fields("M_LOGO") <> "RP" And Not (IsEmpty(proxFicha)) Then
      '    w_desc = "Pg. Saldo Dev.: " & Format(wTXT_TOTAL, "R$ 0.00")
      '    de.cmdIncluirDescCalcVistado Date, proxFicha, 14, "-", wTXT_TOTAL, w_desc, "", "0", "0", "0", adoFichasSaldo.Fields("M_F_COD")
      '    qtdeSaldoAdded = qtdeSaldoAdded + 1
      'End If
          
      'adoFichasSaldo.MoveNext
    'Loop
  
         de.rsTAB_FICHA_MENS.Close
         de.TAB_FICHA_MENS
         
         MsgBox w_reg & " Fichas de funcionários criadas com sucesso!", vbInformation
         w_reg = 0
                   
        Else
            MsgBox "Preencha os Campos!", vbCritical
        End If
    
ElseIf V_RESP = vbYes Then
    MsgBox "Já foram criadas as fichas nesta data!", vbCritical
End If
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
    
End Sub


Private Sub Form_Load()
    TXT_MES = CDbl(Format(Date, "mm"))
    TXT_ANO = Format(Date, "yyyy")
End Sub

Private Sub TXT_ANO_GotFocus()
    Sendkeys "{home}+{end}"
End Sub

Private Sub TXT_ANO_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then Sendkeys "{tab}"

End Sub

Private Sub TXT_LOGO_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then Sendkeys "{tab}"

End Sub

Private Sub TXT_MES_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then Sendkeys "{tab}"

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
