VERSION 5.00
Begin VB.MDIForm mdiPrincipal 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Fichas de Funcionários [SisFF]"
   ClientHeight    =   10755
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   20370
   Icon            =   "mdiPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Menu mnuSis 
      Caption         =   "&Sistema"
      Begin VB.Menu mnuSisCad 
         Caption         =   "Cadastros"
         Begin VB.Menu mnuSisCadFun 
            Caption         =   "Emp"
            Begin VB.Menu mnuSisCadFunAlt 
               Caption         =   "Alteração"
            End
            Begin VB.Menu mnuSisCadFunInc 
               Caption         =   "Incluir"
            End
         End
         Begin VB.Menu mnuSisCadTpC 
            Caption         =   "Tipo de Conta"
            Begin VB.Menu mnuSisCadTpCAlt 
               Caption         =   "Alteração"
            End
            Begin VB.Menu mnuSisCadTpCInc 
               Caption         =   "Incluir"
            End
         End
         Begin VB.Menu mnuSisCadLog 
            Caption         =   "Logo"
            Visible         =   0   'False
            Begin VB.Menu mnuSisCadLogAlt 
               Caption         =   "Alteração"
            End
            Begin VB.Menu mnuSisCadLogInc 
               Caption         =   "Incluir"
            End
         End
         Begin VB.Menu mnuSisVendas 
            Caption         =   "&Lançamentos"
         End
         Begin VB.Menu mnuSisCadSalF 
            Caption         =   "Sal. Família"
         End
         Begin VB.Menu mnuAcessoEspecial 
            Caption         =   "Acesso Especial"
         End
      End
      Begin VB.Menu mnuSisMen 
         Caption         =   "Mensal"
         Begin VB.Menu mnuSisMenFic 
            Caption         =   "Fichas"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSisMenFM 
            Caption         =   "Cadastrar Ficha"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSisMenFicVis 
            Caption         =   "Visualizar Fichas"
         End
         Begin VB.Menu mnuSisMenAFM 
            Caption         =   "Alterar Ficha"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSisMenGer 
            Caption         =   "Gerar Fichas (Todos) *"
         End
         Begin VB.Menu mnuSisMensalVendas 
            Caption         =   "Puxar Vendas"
         End
         Begin VB.Menu mnuSisMenSep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSisMenVisVal 
            Caption         =   "Vistar Contas"
         End
         Begin VB.Menu mnusep04 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnuMaster 
         Caption         =   "Relatórios"
         Begin VB.Menu mnuSisMenFicQtde 
            Caption         =   "Qtde de Emp. por Logo"
         End
         Begin VB.Menu mnuSisMenFicComp 
            Caption         =   "Comparativo de Alterações"
         End
         Begin VB.Menu mnuSisMenFicImpResAn 
            Caption         =   "Resumo Logos - Análitico"
         End
         Begin VB.Menu mnuSisMenFicImpRes 
            Caption         =   "Resumo Logos - Sintético"
         End
         Begin VB.Menu mnuSisMenFicImpTP 
            Caption         =   "Resumo T.P"
         End
         Begin VB.Menu mnuSisMenFicRptEmp 
            Caption         =   "Rel. Emprestimo"
         End
         Begin VB.Menu mnuSisMenFicRptEmpAnalise 
            Caption         =   "Rel. Empréstimo Análise"
         End
         Begin VB.Menu mnuSisMenFicRelVend 
            Caption         =   "Rel. Vendas/Com/Premio"
         End
         Begin VB.Menu mnuSisMenSep10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSisMenFicRelSalCx 
            Caption         =   "Rel. Salarios dos Cxs"
         End
         Begin VB.Menu mnuSep08 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSisMenImp 
            Caption         =   "Imprimir"
            Visible         =   0   'False
            Begin VB.Menu mnuSisMenFicImp 
               Caption         =   "Ficha"
            End
            Begin VB.Menu mnuSisMenFicImpT 
               Caption         =   "Tripa"
            End
         End
      End
      Begin VB.Menu mnuSisSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSisBkp 
         Caption         =   "Backup"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep06 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSisSai 
         Caption         =   "&Sair do Sistema"
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "&Login"
      Begin VB.Menu mnuLogAlt 
         Caption         =   "Alterar"
      End
   End
End
Attribute VB_Name = "mdiPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w_Width_Tela As Integer
Dim w_Height_Tela As Integer
Dim w_Top As Integer
Dim w_Left As Integer
Dim v_Resize As Boolean
Dim w_WindowState As Byte
Dim w_Aberto As Boolean

Dim w_Pic As Byte


'Testa a resolução se for diferente de 800x600 , deixa a janela tamanho restaurado!
Private Sub MDIForm_Activate()



End Sub




Private Sub MDIForm_Load()
    'Me.WindowState = vbMaximized
    
  If v_Resize = False Then
        
        w_Width_Tela = Me.Width
        w_Height_Tela = Me.Height
        
        'Se resulução 1024 x 768 ***
        If Me.Width > 12500 Then
            Me.WindowState = vbNormal
            w_WindowState = WindowState
            'centraliza na tela
            Me.Left = (w_Width_Tela / 2) - (Me.Width / 2)
            Me.Top = (w_Height_Tela / 2) - (Me.Height / 2) - 100
            
            w_WindowState = Me.WindowState
            w_Width_Tela = Me.Width
            w_Height_Tela = Me.Height
            w_Top = Top
            w_Left = Left
            w_Max = False
        Else
            Me.WindowState = vbMaximized
            w_WindowState = WindowState
            w_Max = True
        End If
            
        v_Resize = True
            
    End If
    
   frmMenu.Show
   
    'frm_Alt_Fic_Mensal_VIS.Show
    
sair:
    Exit Sub
err1:
    'MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
    
End Sub

Private Sub MDIForm_Resize()
    
    
    If Me.WindowState <> w_WindowState And WindowState <> 1 And v_Resize = True Or (WindowState <> 1 And Me.WindowState <> w_WindowState And v_Resize = True) Then
        Me.WindowState = w_WindowState
        Me.Width = w_Width_Tela
        Me.Height = w_Height_Tela
        Top = w_Top
        Left = w_Left
    End If
        
    

End Sub



Private Sub mnuAcessoEspecial_Click()
    frm_Alt_Acesso_Especial.Show 1
End Sub

Private Sub mnuImport_Click()
    frm_Import.Show 1
End Sub

Private Sub mnuLogAlt_Click()
   On Error Resume Next
    frm_Cad_Login.Show 1
End Sub

Sub mnuSisBkp_Click()
    Backup
End Sub

Private Sub mnuSisCadFunAlt_Click()
    frm_Alt_Funcionario.Show 1
End Sub
Private Sub mnuSisCadFunInc_Click()
    frm_Cad_Funcionario.Show 1
End Sub


Private Sub mnuSisCadLogAlt_Click()
    frm_Alt_Logo.Show 1
End Sub
Private Sub mnuSisCadLogInc_Click()
    frm_Cad_LOGO.Show 1
End Sub

Private Sub mnuSisCadSalF_Click()
    frm_Cad_Sal_Familia.Show 1
End Sub

Private Sub mnuSisCadTpCAlt_Click()
    frm_Alt_TP_CONTA.Show 1
End Sub
Private Sub mnuSisCadTpCInc_Click()
    frm_Cad_Tp_Conta.Show 1
End Sub
Private Sub mnuSisMenAFM_Click()
    On Error Resume Next
    de.rsTAB_FICHA_MENS.Requery
    If de.rsTAB_FICHA_MENS.RecordCount > 0 Then
        frm_Alt_Fic_Mensal.Show 1
    Else
        MsgBox "Não existe ficha cadastrada!", vbInformation
    End If
End Sub

Private Sub mnuSisMenFicComp_Click()
    frm_Escolha_Comp.Show 1
End Sub

Private Sub mnuSisMenFicImp_Click()
On Error GoTo err1

 FRM_IMP_F.Show 1
 
    
w_mes = FRM_IMP_F.TXT_MES
w_ano = FRM_IMP_F.TXT_ANO
w_Nome = FRM_IMP_F.dbNome & "%"
w_logo = FRM_IMP_F.TXT_LOGO & "%"
    
If FRM_IMP_F.txt_State = "A" And IsNumeric(w_mes) And IsNumeric(w_ano) Then

    If de.rscmdRelFichaMensal_CALC.State = 1 Then de.rscmdRelFichaMensal_CALC.Close
    If de.rscmdRelFichaMensal.State = 1 Then de.rscmdRelFichaMensal.Close
    
    de.cmdRelFichaMensal_CALC w_mes, w_ano, w_Nome, w_logo

    de.cmdRelFichaMensal w_mes, w_ano, w_Nome, w_logo
    
    If Not de.rscmdRelFichaMensal.EOF Then
        de.rscmdRelFichaMensal.MoveFirst
        rptFichaMensal.Show 1
    Else
        MsgBox "NÃO EXISTE FICHAS NO PERÍODO : " & w_mes & "/" & w_ano, vbInformation
    End If
        
    W_CONT = 0
Else
    MsgBox "Relatório Cancelado!", vbInformation
End If
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub mnuSisMenFicImpRes_Click()
On Error GoTo err1
 
    FRM_IMP_F.dbNome.Visible = False
    FRM_IMP_F.ck_Nome.Visible = False
    FRM_IMP_F.lbNome.Visible = False
    FRM_IMP_F.Show 1
    
     
    w_mes = FRM_IMP_F.TXT_MES
    w_ano = FRM_IMP_F.TXT_ANO
    w_Nome = FRM_IMP_F.dbNome
    w_logo = FRM_IMP_F.TXT_LOGO
     
     
    If de.rscmdSqlResumoContasLgSINT.State = 1 Then de.rscmdSqlResumoContasLgSINT.Close
    'de.cmdSqlResumoContasLgSINT w_mes, w_Ano, w_logo
    
    w_Sql = "SELECT TAB_FICHA_MENS.M_MES, TAB_FICHA_MENS.M_ANO, TAB_FICHA_MENS.M_LOGO, TAB_LIQ.Total  " & _
            "AS LIQ, TAB_BRUTO.Total as BRUTO, BRUTO + LIQ AS SALDO " & _
            "FROM TAB_FICHA_MENS, TAB_DESC_CALC, " & _
            "(SELECT TAB_FICHA_MENS.M_MES AS mes , TAB_FICHA_MENS.M_ANO AS ano , TAB_FICHA_MENS.M_LOGO AS logo ,  " & _
            "SUM(TAB_DESC_CALC.C_VALOR) AS Total FROM TAB_FICHA_MENS , TAB_DESC_CALC WHERE TAB_FICHA_MENS.M_NFICHA =  " & _
            "TAB_DESC_CALC.C_N_FICHA AND (TAB_DESC_CALC.C_TP_OP <> '=') AND (TAB_FICHA_MENS.M_BLOQ = 0) and (TAB_DESC_CALC.C_VALOR < 0)   " & _
            "GROUP BY TAB_FICHA_MENS.M_MES , TAB_FICHA_MENS.M_ANO , TAB_FICHA_MENS.M_LOGO) TAB_LIQ ,   " & _
            "(SELECT TAB_FICHA_MENS.M_MES AS mes , TAB_FICHA_MENS.M_ANO AS ano , TAB_FICHA_MENS.M_LOGO AS logo ,  " & _
            "SUM(TAB_DESC_CALC.C_VALOR) AS Total FROM TAB_FICHA_MENS , TAB_DESC_CALC WHERE TAB_FICHA_MENS.M_NFICHA =  " & _
            "TAB_DESC_CALC.C_N_FICHA AND (TAB_DESC_CALC.C_TP_OP <> '=') AND (TAB_FICHA_MENS.M_BLOQ = 0) and (TAB_DESC_CALC.C_VALOR > 0)  " & _
            "GROUP BY TAB_FICHA_MENS.M_MES , TAB_FICHA_MENS.M_ANO , TAB_FICHA_MENS.M_LOGO) TAB_BRUTO  " & _
            "Where TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA  " & _
            "And (TAB_FICHA_MENS.M_MES = TAB_LIQ.mes AND TAB_FICHA_MENS.M_MES = TAB_BRUTO.mes) " & _
            "AND (TAB_FICHA_MENS.M_ANO = TAB_LIQ.ano AND TAB_FICHA_MENS.M_ANO = TAB_BRUTO.ano) " & _
            "AND (TAB_FICHA_MENS.M_LOGO = TAB_LIQ.logo AND TAB_FICHA_MENS.M_LOGO = TAB_BRUTO.logo) " & _
            "AND (TAB_DESC_CALC.C_TP_OP <> '=')  " & _
            "GROUP BY TAB_FICHA_MENS.M_MES, TAB_FICHA_MENS.M_ANO, TAB_FICHA_MENS.M_LOGO, TAB_LIQ.Total,TAB_BRUTO.Total " & _
            "HAVING (TAB_FICHA_MENS.M_MES = " & w_mes & ") AND (TAB_FICHA_MENS.M_ANO = " & w_ano & ") AND  " & _
            "(TAB_FICHA_MENS.M_LOGO LIKE '" & w_logo & "') ORDER BY TAB_FICHA_MENS.M_MES, TAB_FICHA_MENS.M_ANO"

    de.rscmdSqlResumoContasLgSINT.Open w_Sql
    
    rptRelResumoContasLg.Sections(2).Controls("lbPer").Caption = "  Período :  " & Format(w_mes, "00") & " / " & w_ano
    rptRelResumoContasLg.Show 1
    
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub mnuSisMenFicImpResAn_Click()
On Error GoTo err1
 
    FRM_IMP_F.Show 1
     
    w_mes = FRM_IMP_F.TXT_MES
    w_ano = FRM_IMP_F.TXT_ANO
    w_Nome = FRM_IMP_F.dbNome
    w_logo = FRM_IMP_F.TXT_LOGO
    
    If de.rscmdSqlResumoContasLg_Grouping.State = 1 Then de.rscmdSqlResumoContasLg_Grouping.Close
    de.cmdSqlResumoContasLg_Grouping w_mes, w_ano, w_logo
    rptRelResumoContasLgDet.Show 1
    
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume sair

End Sub

Private Sub mnuSisMenFicImpT_Click()
Dim SQL_Tripa As String
Dim SQL_TripaDet As String
Dim SQL_TripaFichas As String

Dim w_Lojas As String
Dim w_sqlLojasTripa As String

Dim w_Tipos As String
Dim w_sqlTiposTripa As String

Dim w_FirstLoja As Boolean
Dim w_FirstFicha As Boolean
Dim w_FirstTipo As Boolean
 
Dim w_tipoTripa(50) As Variant
Dim w_lojaTripa(100) As Variant
 
On Error GoTo err1
     
    
    If FRM_IMP_F.ckTodas.value = 0 Then
        FRM_IMP_F.TXT_LOGO = txtLogo
    Else
        FRM_IMP_F.TXT_LOGO = "%"
    End If
    
    FRM_IMP_F.TXT_MES = TXT_MES
    FRM_IMP_F.TXT_ANO = TXT_ANO
    
    If FRM_IMP_F.ck_Nome.value = 0 Then
        FRM_IMP_F.dbNome = TXT_FUNC
    Else
        FRM_IMP_F.dbNome = "%"
    End If
    FRM_IMP_F.CkFicha.Visible = True
    FRM_IMP_F.CkTripa.Visible = True
    

    FRM_IMP_F.txt_tipo = TXT_FTIPO
    
    FRM_IMP_F.Show 1
    
    If (FRM_IMP_F.txt_State = "F") Then 'Or (FRM_IMP_F.CkTripa.value = 1 And FRM_IMP_F.CkFicha.value = 1) Then
       MsgBox "Impressão Cancelada!", vbCritical
    Else
        
    'lojas
    w_FirstLoja = True
    For I = 0 To FRM_IMP_F.TXT_LOGO.ListCount - 1
        If FRM_IMP_F.TXT_LOGO.Selected(I) = True Then
            If w_FirstLoja Then
                w_Lojas = "'" & FRM_IMP_F.TXT_LOGO.list(I) & "'"
                w_sqlLojasTripa = " TAB_FUNCIONARIO.F_Cod_L = '" & FRM_IMP_F.TXT_LOGO.list(I) & "' "
            Else
                w_Lojas = w_Lojas & ",'" & FRM_IMP_F.TXT_LOGO.list(I) & "'"
                w_sqlLojasTripa = w_sqlLojasTripa & " OR TAB_FUNCIONARIO.F_Cod_L = '" & FRM_IMP_F.TXT_LOGO.list(I) & "' "
            End If
            w_lojaTripa(I) = FRM_IMP_F.TXT_LOGO.list(I)
            w_FirstLoja = False
        End If
    Next
    
    
    'tipos
    w_FirstTipo = True
    Dim w_tipo
    For J = 0 To FRM_IMP_F.txt_tipo.ListCount - 1
        If FRM_IMP_F.txt_tipo.Selected(J) = True Then
            Select Case FRM_IMP_F.txt_tipo.list(J)
                Case "VENDEDOR": w_tipo = "V"
                Case "GERENTE": w_tipo = "G"
                Case "GER RODA": w_tipo = "D"
                Case "CAIXA": w_tipo = "C"
                Case "2º CAIXA": w_tipo = "2"
                Case "CX EXTRA": w_tipo = "X"
                Case "SEGURANÇA": w_tipo = "R"
                Case "SUPERVISOR": w_tipo = "S"
                Case "RP": w_tipo = "O"
            End Select
        
            If w_FirstTipo Then
                w_Tipos = "'" & w_tipo & "'"
                w_sqlTiposTripa = " TAB_FUNCIONARIO.F_TIPO = '" & w_tipo & "' "
            Else
                w_Tipos = w_Tipos & ",'" & w_tipo & "'"
                w_sqlTiposTripa = w_sqlTiposTripa & " OR TAB_FUNCIONARIO.F_TIPO = '" & w_tipo & "' "
            End If
            w_tipoTripa(J) = w_tipo
            w_FirstTipo = False
        End If
    Next
    
   
        If FRM_IMP_F.CkTripa.value = 1 Then
                
            
            If de.rscmdRelFichaMensal_TRIPA.State = 1 Then de.rscmdRelFichaMensal_TRIPA.Close
            
            'SQL_Tripa = "SELECT TAB_FICHA_MENS.M_NFICHA AS Ficha, TAB_FUNCIONARIO.F_NOME AS Nome," _
            '    & "Format('01/'+Mid(Str(TAB_FICHA_MENS.M_MES),2)+'/'+Mid(Str(TAB_FICHA_MENS.M_ANO),2),'DD/MM/YYYY') AS Data," _
            '    & "TAB_FUNCIONARIO.F_Cod_L AS Logo2, LOJB010.NUM AS Logo, TAB_FICHA_MENS.M_TOTAL, Mid(TAB_FUNCIONARIO.F_COD_CENTRAL,3) AS COD_CENTRAL," _
            '    & " TAB_FUNCIONARIO.F_TIPO AS TIPO, TAB_FUNCIONARIO.F_CX_QT_VND AS Cx_Qt_VND FROM TAB_FICHA_MENS, TAB_FUNCIONARIO, LOJB010" _
            '    & " WHERE (LOJB010.COD_LOJ = TAB_FUNCIONARIO.F_Cod_L) AND (((TAB_FICHA_MENS.M_F_COD)=[TAB_FUNCIONARIO].[F_Codigo]) AND ((TAB_FICHA_MENS.M_MES)=" & FRM_IMP_F.txt_Mes & ") AND" _
            '    & " ((TAB_FICHA_MENS.M_ANO)=" & FRM_IMP_F.txt_Ano & ") AND ((TAB_FUNCIONARIO.F_NOME) Like '" & FRM_IMP_F.dbNome & "' and TAB_FUNCIONARIO.F_NOME <> '10 - Func'" _
            '    & " AND TAB_FUNCIONARIO.F_NOME <> '99 - Presence') AND (" _
            '    & w_sqlTiposTripa _
            '    & ") AND (" _
            '    & w_sqlLojasTripa _
            '    & ")) GROUP BY TAB_FICHA_MENS.M_NFICHA, TAB_FUNCIONARIO.F_NOME," _
            '    & " Format('01/'+Mid(Str(TAB_FICHA_MENS.M_MES),2)+'/'+Mid(Str(TAB_FICHA_MENS.M_ANO),2),'DD/MM/YYYY')," _
            '    & " TAB_FUNCIONARIO.F_Cod_L, TAB_FICHA_MENS.M_TOTAL, Mid(TAB_FUNCIONARIO.F_COD_CENTRAL,3), TAB_FUNCIONARIO.F_TIPO," _
            '    & " TAB_FUNCIONARIO.F_CX_QT_VND, Len(TAB_FICHA_MENS.M_DT_ACF) HAVING (((TAB_FUNCIONARIO.F_NOME) ) AND ((TAB_FUNCIONARIO.F_Cod_L) )" _
            '    & " AND ((Len([TAB_FICHA_MENS].[M_DT_ACF])) IS NULL)) OR (((Len([TAB_FICHA_MENS].[M_DT_ACF]))<5)) " _
            '    & "ORDER BY TAB_FUNCIONARIO.F_Cod_L, TAB_FUNCIONARIO.F_TIPO DESC , TAB_FUNCIONARIO.F_NOME;"
            SQL_Tripa = "SELECT TAB_FICHA_MENS.M_NFICHA AS Ficha, TAB_FUNCIONARIO.F_NOME AS Nome," _
                & "Format('01/'+Mid(Str(TAB_FICHA_MENS.M_MES),2)+'/'+Mid(Str(TAB_FICHA_MENS.M_ANO),2),'DD/MM/YYYY') AS Data," _
                & "TAB_FUNCIONARIO.F_Cod_L AS Logo2, LOJB010.NUM as Logo, TAB_FICHA_MENS.M_TOTAL, Mid(TAB_FUNCIONARIO.F_COD_CENTRAL,3) AS COD_CENTRAL," _
                & " TAB_FICHA_MENS.M_TIPO AS TIPO, TAB_FUNCIONARIO.F_CX_QT_VND AS Cx_Qt_VND FROM TAB_FICHA_MENS, TAB_FUNCIONARIO INNER JOIN Lojb010 ON TAB_FUNCIONARIO.F_Cod_L = Lojb010.COD_LOJ " _
                & " WHERE (((TAB_FICHA_MENS.M_F_COD)=[TAB_FUNCIONARIO].[F_Codigo]) AND ((TAB_FICHA_MENS.M_MES)=" & FRM_IMP_F.TXT_MES & ") AND" _
                & " ((TAB_FICHA_MENS.M_ANO)=" & FRM_IMP_F.TXT_ANO & ") AND ((TAB_FUNCIONARIO.F_NOME) Like '" & FRM_IMP_F.dbNome & "' and TAB_FUNCIONARIO.F_NOME <> '10 - Func'" _
                & " AND TAB_FUNCIONARIO.F_NOME <> '99 - Presence') AND   (" _
                & w_sqlTiposTripa _
                & ") AND (" _
                & w_sqlLojasTripa _
                & ")) GROUP BY TAB_FICHA_MENS.M_NFICHA, TAB_FUNCIONARIO.F_NOME," _
                & " Format('01/'+Mid(Str(TAB_FICHA_MENS.M_MES),2)+'/'+Mid(Str(TAB_FICHA_MENS.M_ANO),2),'DD/MM/YYYY')," _
                & " TAB_FUNCIONARIO.F_Cod_L, TAB_FICHA_MENS.M_TOTAL, Mid(TAB_FUNCIONARIO.F_COD_CENTRAL,3), TAB_FICHA_MENS.M_TIPO," _
                & " TAB_FUNCIONARIO.F_CX_QT_VND, Len(TAB_FICHA_MENS.M_DT_ACF), LOJB010.NUM HAVING (((TAB_FUNCIONARIO.F_NOME) ) AND ((TAB_FUNCIONARIO.F_Cod_L) )" _
                & " AND ((Len([TAB_FICHA_MENS].[M_DT_ACF])) IS NULL)) OR (((Len([TAB_FICHA_MENS].[M_DT_ACF]))<5)) " _
                & "ORDER BY LOJB010.NUM, TAB_FICHA_MENS.M_TIPO DESC , TAB_FUNCIONARIO.F_NOME;"
                
                'TXT_OBS = SQL_Tripa
            de.rscmdSqlTripa.Open SQL_Tripa, , adOpenStatic, adLockOptimistic
            
            
            If Not de.rscmdSqlTripa.EOF Then
                If de.rscmdSqlTotalVND.State = 1 Then de.rscmdSqlTotalVND.Close
                
            
                w_DtI = CVDate("01/" & Format(FRM_IMP_F.TXT_MES, "00") & "/" & Format(FRM_IMP_F.TXT_ANO, "0000"))
                w_DtF = UltDiaMes(FRM_IMP_F.TXT_MES, FRM_IMP_F.TXT_ANO)
                de.cmdSqlTotalVND w_DtI, w_DtF, IIf(FRM_IMP_F.TXT_LOGO = "", "%", FRM_IMP_F.TXT_LOGO)
                
                
                                'fichas
                w_FirstFicha = True
                
                de.rscmdSqlTripa.MoveFirst
                Do While Not de.rscmdSqlTripa.EOF
                    If w_FirstFicha Then
                        SQL_TripaFichas = " TAB_DESC_CALC.C_N_FICHA = " & de.rscmdSqlTripa.Fields("Ficha") & " "
                
                    Else
                        SQL_TripaFichas = SQL_TripaFichas & " OR TAB_DESC_CALC.C_N_FICHA = " & de.rscmdSqlTripa.Fields("Ficha") & " "
                    End If
                    w_FirstFicha = False
                    de.rscmdSqlTripa.MoveNext
                Loop

                de.rscmdSqlTripa.MoveFirst
                
                If de.rscmdSqlTripaDet.State = 1 Then de.rscmdSqlTripaDet.Close
                SQL_TripaDet = "SELECT TAB_DESC_CALC.C_N_FICHA AS Ficha, TAB_DESC_CALC.C_TP_CONTA, TAB_TP_CONTA.TP_DESC AS Conta," _
                    & "SUM(TAB_DESC_CALC.C_VALOR) AS Valor, TAB_DESC_CALC.C_TP_OP AS Op, TAB_TP_CONTA.TP_NIVEL FROM TAB_FICHA_MENS," _
                    & "TAB_DESC_CALC, TAB_TP_CONTA WHERE TAB_FICHA_MENS.M_NFICHA = TAB_DESC_CALC.C_N_FICHA " _
                    & "AND TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (" _
                    & SQL_TripaFichas _
                    & ") GROUP BY TAB_DESC_CALC.C_N_FICHA, TAB_TP_CONTA.TP_DESC," _
                    & "TAB_DESC_CALC.C_TP_OP, TAB_DESC_CALC.C_TP_CONTA, TAB_TP_CONTA.TP_NIVEL ORDER BY TAB_DESC_CALC.C_N_FICHA," _
                    & "SUM(TAB_DESC_CALC.C_VALOR) DESC"
                
                de.rscmdSqlTripaDet.Open SQL_TripaDet, , adOpenStatic, adLockOptimistic
                
                
                Set AdoItem1 = de.rscmdSqlTripaDet.Clone
                'Criar_RPT_TRIPA de.rscmdRelFichaMensal_TRIPA, AdoItem1
                PrintTripa de.rscmdSqlTripa, AdoItem1, (FRM_IMP_F.ck_Nome.value = 0 And FRM_IMP_F.ckTodas.value = 0)
                frmTripa.Show 1
            End If
        End If
        End If

    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume sair

End Sub


Private Sub mnuSisMenFicImpTP_Click()
On Error GoTo err1

 FRM_IMP_F.dbNome.Visible = False
 FRM_IMP_F.lbNome.Visible = False
 FRM_IMP_F.ck_Nome.Visible = False
 
 FRM_IMP_F.Show 1
 
w_mes = FRM_IMP_F.TXT_MES
w_ano = FRM_IMP_F.TXT_ANO
w_logo = FRM_IMP_F.TXT_LOGO & "%"
    
If FRM_IMP_F.txt_State = "A" And IsNumeric(w_mes) And IsNumeric(w_ano) Then
    
    If de.rscmdSqlTP.State = 1 Then de.rscmdSqlTP.Close
    de.cmdSqlTP w_mes, w_ano, w_logo
    
    If Not de.rscmdSqlTP.EOF Then
        
        rptRelTP.Sections("seccab").Controls("lbPer").Caption = "  Período :  " & Format(w_mes, "00") & " / " & w_ano
        
        wTot = 0
        
        Do While Not de.rscmdSqlTP.EOF
            wTot = wTot + CDbl(de.rscmdSqlTP.Fields("TOTAL_TP"))
            de.rscmdSqlTP.MoveNext
        Loop
        
        rptRelTP.Sections("secrod").Controls("lbTot").Caption = Format(wTot, "0")
        rptRelTP.Show
        
    Else
        MsgBox "NÃO EXISTE T.P NESTE PERÍODO : " & w_mes & "/" & w_ano, vbInformation
    End If
        
    W_CONT = 0
Else
    MsgBox "Relatório Cancelado!", vbInformation
End If
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume sair

End Sub

Private Sub mnuSisMenFicQtde_Click()
On Error GoTo err1

     If de.rscmdQtde_Func_Logo_Grouping.State = 1 Then de.rscmdQtde_Func_Logo_Grouping.Close
ini:

     w_mes = InputBox("Entre com o Mês:", , Format(Date, "MM"))
     w_ano = InputBox("Entre com o Ano:", , Format(Date, "YYYY"))
     If IsNumeric(w_mes) And IsNumeric(w_ano) Then
     
         de.cmdQtde_Func_Logo_Grouping w_ano, w_mes
         rptRelQtdeEmp.Sections("SecCab").Controls.Item("LbPer").Caption = "Período : " & w_mes & "/" & w_ano
         
         If UCase(InputBox("Mostrar Nomes ?" & Chr(13) & Chr(13) & "S - Sim" & Chr(13) & "N - Não", "Opção", "S")) = "N" Then
              rptRelQtdeEmp.Sections("SecDet").Visible = False
              rptRelQtdeEmp.Sections("SecCabG").Controls.Item("Fundo").BackColor = &HFFFFFF
              rptRelQtdeEmp.Sections("SecCabG").Controls.Item("LB1").Visible = False
              rptRelQtdeEmp.Sections("SecCabG").Controls.Item("LB2").Visible = False
              rptRelQtdeEmp.Sections("SecCabG").Controls.Item("LB3").Visible = False
              rptRelQtdeEmp.Sections("SecCabG").Controls.Item("LB4").Visible = False
         
         End If
         rptRelQtdeEmp.Show
     
     Else
        MsgBox "Redigite o Mês e Ano Desejado!", vbExclamation
        GoTo ini
     End If


sair:
    Exit Sub
err1:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub mnuSisMenFicRelSalCx_Click()
On Error GoTo err1

     If de.rscmdSqlSalarioCX_Grouping.State = 1 Then de.rscmdSqlSalarioCX_Grouping.Close
ini:

     w_mes = InputBox("Entre com o Mês:", , Format(Date, "MM"))
     w_ano = InputBox("Entre com o Ano:", , Format(Date, "YYYY"))
     If IsNumeric(w_mes) And IsNumeric(w_ano) Then
     
         de.cmdSqlSalarioCX_Grouping w_mes, w_ano
         
         rptSalarioCx.Show
     
     Else
        MsgBox "Redigite o Mês e Ano Desejado!", vbExclamation
        GoTo ini
     End If


sair:
    Exit Sub
err1:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume sair

End Sub

Private Sub mnuSisMenFicRelVend_Click()
    FRM_IMP_F.Show 1
     
    'w_mes = FRM_IMP_F.TXT_MES
    'w_ano = FRM_IMP_F.TXT_ANO
    'w_Nome = FRM_IMP_F.dbNome
    'w_logo = FRM_IMP_F.TXT_LOGO
      
    de.sqlComissaoPremio FRM_IMP_F.TXT_MES, FRM_IMP_F.TXT_ANO, FRM_IMP_F.TXT_LOGO
    rptComissaoPremio.Show
    
        
    
    
    
    'FRM_IMP_F.Show 1
     
    'w_mes = FRM_IMP_F.TXT_MES
    'w_ano = FRM_IMP_F.TXT_ANO
    'w_Nome = FRM_IMP_F.dbNome
    'w_logo = FRM_IMP_F.TXT_LOGO
    
    
    'wSQL = " SHAPE {SELECT * FROM `Con_Rpt_Com_Vendas` " & _
           " WHERE (M_LOGO LIKE '" & w_logo & "') AND (M_MES = " & w_mes & ") AND (M_ANO = " & w_ano & ") AND (F_NOME LIKE '" & w_Nome & "')" & _
           "}  AS Con_Rpt_Com_Vendas COMPUTE Con_Rpt_Com_Vendas BY 'M_LOGO'"
    
    'If de.rsCon_Rpt_Com_Vendas_Grouping.State = 1 Then de.rsCon_Rpt_Com_Vendas_Grouping.Close
    'de.rsCon_Rpt_Com_Vendas_Grouping.Open wSQL

    'rptVendasCom.Sections("Cab").Controls("lbTitulo").Caption = " Ref.  " & w_mes & "/" & w_ano
    'rptVendasCom.Show
End Sub

Private Sub mnuSisMenFicRptEmp_Click()
    rptEmprestimos.Show
End Sub

Private Sub mnuSisMenFicRptEmpAnalise_Click()
    rptEmprestimoAnalise.Show
End Sub

Private Sub mnuSisMenFicVis_Click()
    On Error Resume Next

    frm_Alt_Fic_Mensal_VIS.Show '1

End Sub
Private Sub mnuSisMenFicVis_Click2()
    On Error Resume Next
    
    de.rsTAB_FICHA_MENS.Requery
    
    
    frm_Alt_Fic_Mensal_Visualizar.txt_PMes = InputBox("Entre com o Mês:", , Format(Date, "MM"))
    frm_Alt_Fic_Mensal_Visualizar.txt_PAno = InputBox("Entre com o Ano:", , Format(Date, "YYYY"))
    
    If de.rscmdSqlVisualizarFichas.State = 1 Then de.rscmdSqlVisualizarFichas.Close
    de.cmdSqlVisualizarFichas frm_Alt_Fic_Mensal_Visualizar.txt_PAno, frm_Alt_Fic_Mensal_Visualizar.txt_PMes
    
    If de.rscmdSqlVisualizarFichas.RecordCount > 0 Then
        frm_Alt_Fic_Mensal_Visualizar.Show 1
    Else
        MsgBox "Não existe ficha cadastrada!", vbInformation
    End If
    'frm_Alt_Fic_Mensal_VIS.Show 1

End Sub

Private Sub mnuSisMenFM_Click()
    frm_Cad_Fic_Mensal.Show 1
End Sub
Private Sub mnuSisMenGer_Click()
    frm_Gerar_Fichas.Show 1
End Sub


Private Sub mnuSisMensalVendas_Click()
Dim w_Data_server

    Call Put_System_CNC
    Call AbreConexao(Conexão, "")
    w_Data_server = ExecuteSQL("SELECT COUNT(*) AS TOTAL FROM FILIAIS", , , False).Fields(0)
    MsgBox w_Data_server

End Sub

Private Sub mnuSisMenVisVal_Click()
    frm_Alt_Visto_Vale.Show 1
End Sub

Private Sub mnuSisSai_Click()
    If vbYes = MsgBox("Deseja realmente Sair?", vbQuestion + vbYesNo) Then
        End
    Else
        frmMenu.Show
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    mnuSisSai_Click
    
    Cancel = -1
End Sub

Private Sub mnuSisVendas_Click()
    frm_Vendas.Show
End Sub
