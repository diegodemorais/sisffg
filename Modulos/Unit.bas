Attribute VB_Name = "Unit"
Public md5Test As MD5

Public strConectaFirebird As String

Public w_AcessoEspecial As Boolean

Public strDirBase As String  'define o diretório do Banco de dados
Public strDirBaseCentral As String  'define o diretório do Banco de dados da Central
Public strDirBaseServer As String  'define o diretório do Banco de dados do Servidor
Public strDirRPT As String  ' define o diretório dos relatorios
Public strImpressora As String 'Define o local de impressão da tripa
Public strImgFundo As String 'define Imagem de fundo
Public strImgSplash As String 'define Imagem de fundo
Public W_CONT  As Byte 'CONTA AS VEZES Q/ ABRIU O RELATORIO

Public p_Usuario As String

Public w_CodFunc As String

Public w_umaVez As Byte

Public w_Func_atual As String

Dim w_pg_vt As Boolean

'Não permite editar os lançamentos (frm_alt_desc_calc)
Public w_leitura As Boolean

'*** VARIAVEIS P/ CALCULOS DE JUROS DO EMPRESTIMOS ***
Public W_JURO_AO_MES As Double
Public W_JURO_AO_DIA As Double
Public W_DT_ULT_PG As Date
Public W_PARC_RESTANTE As Byte
Public W_SALDO As Currency
Public w_qt_dias As Integer
Public W_QT_MESES As Double
Public W_SALDO_AC As Currency

Public w_Max As Boolean  '*  se tela esta maximizada

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As String, ByVal lpDefault As String, ByVal _
lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName _
As String) As Long

Public DBs As dao.Database
Public wTCli As dao.Recordset
Public w_usuario, w_usuario2

'*** variavel de DAO p/ Func  Cred
Public db As dao.Database
Public wtabFuncCred As dao.Recordset

Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

'*** CONSTANTES
    '*** pl ***
    Public Const NomeMestre = "PL"
    Public Const NomeMestre2 = "DIEGO"
    Public Const NomeMestre3 = "FLAVIO"
    
    '*** USUario ***
    Public Const NomeUsu = "BEL"
    Public Const SenhaUsu = "74123"
    
    Public Const NomeUsu2 = "KELEN"
    Public Const SenhaUsu2 = "kma1403"
    
    Public Const NomeUsu3 = "RODRIGO"
    Public Const SenhaUsu3 = "547985"
    
    Public Const NomeUsu4 = "FLAVIO2"
    Public Const SenhaUsu4 = "625731"
    
    Public Const NomeUsu5 = "KELY"
    Public Const SenhaUsu5 = "2703"
    
    Public w_PassWordLib As String
    Public SenhaMestre As String

'O problema do FileCopy do VB é q ele não mostra visualmente a
'operação (barra de progresso e etc) com no Explorer. Ao invés
'de copiar arquivos com o FileCopy, use a rotina API abaixo:

'Num módulo:

Public Declare Function SHFileOperation Lib "shell32.dll" _
       Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) _
       As Long

Public Const FO_COPY As Long = &H2
Public Const FOF_ALLOWUNDO As Long = &H40

Public Type SHFILEOPSTRUCT
  hwnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Boolean
  hNameMappings As Long
  lpszProgressTitle As String
End Type


Public Type bVal
    b As String
    val As Double
End Type

Public Sub CopiarArq(Origem As String, Destino As String)
  Dim RST As Long
  Dim FLOP As SHFILEOPSTRUCT

  FLOP.hwnd = 0
  FLOP.wFunc = FO_COPY

  'Arquivo de origem:
  FLOP.pFrom = Origem & vbNullChar & vbNullChar

  'Para copiar TODOS os arquivos, use:
  'FLOP.pFrom = "C:\*.*" & vbNullChar & vbNullChar

  'Diretório ou arquivo de destino:
  FLOP.pTo = Destino & vbNullChar & vbNullChar

  FLOP.fFlags = FOF_ALLOWUNDO
  RST = SHFileOperation(FLOP)
  If RST <> 0 Then
    'Erro na cópia
    'MsgBox Err.LastDllError, vbCritical Or vbOKOnly
  Else
    If FLOP.fAnyOperationsAborted <> 0 Then
      'MsgBox "Falha na cópia!!!", vbCritical Or vbOKOnly
    End If
  End If
End Sub

Function isMesValido(cod_func As Variant, mes As Variant, ano As Variant) As Boolean
    Dim mes_atual, ano_atual, mes_ant, ano_ant

    ano_atual = de.cnc.Execute("SELECT Max([M_ANO]) FROM TAB_FUNCIONARIO INNER JOIN TAB_FICHA_MENS ON TAB_FUNCIONARIO.F_Codigo = TAB_FICHA_MENS.M_F_COD WHERE TAB_FUNCIONARIO.F_Codigo= " & cod_func & "").Fields(0)
    mes_atual = de.cnc.Execute("SELECT Max([M_MES]) FROM TAB_FUNCIONARIO INNER JOIN TAB_FICHA_MENS ON TAB_FUNCIONARIO.F_Codigo = TAB_FICHA_MENS.M_F_COD WHERE TAB_FUNCIONARIO.F_Codigo= " & cod_func & " AND TAB_FICHA_MENS.M_ANO = " & ano_atual).Fields(0)

    'mes_ant = mes_atual - 1
    'If mes_ant = 0 Then
        'mes_ant = 12
        'ano_ant = ano_atual - 1
   ' Else
        'ano_ant = ano_atual
    'End If
    
    If CInt(mes) >= mes_atual Then
        If CInt(ano) >= ano_atual Then
            isMesValido = True
        Else
            isMesValido = False
        End If
    Else
        isMesValido = False
    End If
    
    'If (Format(mes, "##") = Format(mes_atual, "##")) Or (Format(mes, "##") = Format(mes_ant, "##")) Then
        'If (ano = ano_atual) Or (ano = ano_ant) Then
            'isMesValido = True
        'Else
            'isMesValido = False
        'End If
    'Else
        'isMesValido = False
    'End If
    
End Function

Function isMesValido2(cod_func As Variant, mes As Variant, ano As Variant) As Boolean
    Dim mes_atual, ano_atual, mes_ant, ano_ant

    ano_atual = de.cnc.Execute("SELECT Max([M_ANO]) FROM TAB_FUNCIONARIO INNER JOIN TAB_FICHA_MENS ON TAB_FUNCIONARIO.F_Codigo = TAB_FICHA_MENS.M_F_COD WHERE TAB_FUNCIONARIO.F_Codigo= " & cod_func & "").Fields(0)
    mes_atual = de.cnc.Execute("SELECT Max([M_MES]) FROM TAB_FUNCIONARIO INNER JOIN TAB_FICHA_MENS ON TAB_FUNCIONARIO.F_Codigo = TAB_FICHA_MENS.M_F_COD WHERE TAB_FUNCIONARIO.F_Codigo= " & cod_func & " AND TAB_FICHA_MENS.M_ANO = " & ano_atual).Fields(0)

    mes_ant = mes_atual - 1
    If mes_ant = 0 Then
        mes_ant = 12
        ano_ant = ano_atual - 1
    Else
        ano_ant = ano_atual
    End If
    
     
    If CInt(mes) >= mes_ant Then
        If CInt(ano) >= ano_atual Or CInt(ano) = ano_ant Then
            isMesValido2 = True
        Else
            isMesValido2 = False
        End If
    Else
        isMesValido2 = False
    End If
    
End Function

Sub Pause(Seconds As Single)
Dim EndTime As Date

    EndTime = DateAdd("s", Seconds, Now)
    
    Do
    DoEvents
    Loop Until Now >= EndTime

End Sub


Sub Desligar()
Select Case MsgBox("Deseja Desligar o Computador?", vbInformation + vbYesNoCancel + vbDefaultButton2)
Case 6
    Call ExitWindowsEx(1, 1)
Case 7
    End
End Select

End Sub

Public Function DCount(RsOrigem As ADODB.Recordset, Criterio As String) As Long
On Error GoTo err1
    'função de pesquisa num record set
    Dim rs As ADODB.Recordset
    Set rs = RsOrigem
    If rs.State = 0 Then
     rs.Open
    End If
    rs.Filter = Criterio
    If rs.EOF <> True Then
       DCount = rs.RecordCount  'achou
    Else
       DCount = rs.RecordCount  'nào achou
    End If
        
sair:
    
    Exit Function
err1:
    MsgBox "Erro no DCount - pesquisa ou critério" & Chr(13) & Chr(13) & "Err: " & Error$, vbCritical
    Resume sair
    
End Function
 
Public Function Dlookup(RsOrigem As ADODB.Recordset, Criterio As String) As ADODB.Recordset
On Error GoTo err1
    'função de pesquisa num record set
    Dim rs As ADODB.Recordset
    Set rs = RsOrigem
    rs.Filter = Criterio
    If rs.EOF <> True Then
       Set Dlookup = rs 'achou
    Else
       Set Dlookup = rs 'nào achou
    End If
        
sair:
    
    Exit Function
err1:
    MsgBox "Erro no Dlookup - pesquisa ou critério" & Chr(13) & Chr(13) & "Err: " & Error$, vbCritical
    Resume sair
    
End Function

Public Sub KeyEnter(key)
    'Utilizar assim : KeyEnter (KeyAscii)
    Select Case key
    Case 13 'enter
        Sendkeys "{tab}{home}+{end}"
    Case 38 'seta para baixo
        Sendkeys "+{tab}{home}+{end}"
    Case 40 'seta para cima
        Sendkeys "{tab}{home}+{end}"
    End Select
End Sub



Public Function DiaSemana(Num As Integer, Abreviado As Boolean) As String
        Select Case Num 'Verifica Nº do dia da semana para escrever o dia referente
          Case 1:
                 If Abreviado = True Then
                    DiaSemana = "Dom"
                 Else
                    DiaSemana = "Domingo"
                 End If
          Case 2:
                 If Abreviado = True Then
                    DiaSemana = "Seg"
                 Else
                    DiaSemana = "2ª Feira"
                 End If
          Case 3:
                 If Abreviado = True Then
                    DiaSemana = "Ter"
                 Else
                    DiaSemana = "3ª Feira"
                 End If
          Case 4:
                 If Abreviado = True Then
                    DiaSemana = "Qua"
                 Else
                    DiaSemana = "4ª Feira"
                 End If
          Case 5:
                 If Abreviado = True Then
                    DiaSemana = "Qui"
                 Else
                    DiaSemana = "5ª Feira"
                 End If
          Case 6:
                 If Abreviado = True Then
                    DiaSemana = "Sex"
                 Else
                    DiaSemana = "6ª Feira"
                 End If
          Case 7:
                 If Abreviado = True Then
                    DiaSemana = "Sab"
                 Else
                    DiaSemana = "Sábado"
                 End If

         End Select
End Function





Public Sub AtualizarGeral()
On Error GoTo err1
'    Static w_umaVez As Byte
     
If w_umaVez = 0 Then
       
    
    frmSplash.PB.value = 5
    frmSplash.PB.text = "Carregando " & frmSplash.PB.value & "%"

    'ABRE AS TABELAS
       'MSAccess
       de.TAB_DESC_CALC
       de.TAB_FICHA_MENS
       de.TAB_TP_CONTA
       de.TAB_FUNCIONARIO
       
       
    frmSplash.PB.value = 15
    frmSplash.PB.text = "Carregando " & frmSplash.PB.value & "%"
       
       de.TAB_FUNC_CENTRAL
       
       
    frmSplash.PB.value = 20
    frmSplash.PB.text = "Carregando " & frmSplash.PB.value & "%"
       'Dbase
       de.TAB_L
    frmSplash.PB.value = 30
    frmSplash.PB.text = "Carregando " & frmSplash.PB.value & "%"


       'Exclui todos os Registro de Tab_AUX
       de.cmdExcluirAuxCred

    frmSplash.PB.value = 35
    frmSplash.PB.text = "Carregando " & frmSplash.PB.value & "%"
  
  
    frmSplash.PB.value = 62
    frmSplash.PB.text = "Carregando " & frmSplash.PB.value & "%"
  
    'Paradox
    'abri os Clientes de Crediarios
    'If frmLogin.Option1.value = True Then de.TAB_FUNC_CRED
  
    frmSplash.PB.value = 84
    frmSplash.PB.text = "Carregando " & frmSplash.PB.value & "%"
  
    Pause 0.3
    frmSplash.PB.value = 100
    frmSplash.PB.text = "Carregando " & frmSplash.PB.value & "%"
    
    Pause 0.5
       
       w_umaVez = 1
End If

sair:
    Exit Sub
err1:
    
    If (UCase(txtUserName.text) = NomeMestre Or UCase(txtUserName.text) = NomeMestre2 Or UCase(txtUserName.text) = NomeMestre3) Then
        
        If CDbl(Err.Number) = CDbl(424) Then
            'MsgBox "Favor fazer a importação de tabelas!", vbCritical
        Else
             MsgBox Error$ & dePP.Recordsets(I).Source, vbCritical
        End If
    End If
    Resume sair
End Sub

Private Sub ACRESCENTA_BARRA()
Dim Porcentagem As Integer

On Error GoTo err1
    'frmSplash.pgbEntrada.Value = frmSplash.pgbEntrada.Value + 1
    ''Porcentagem = (frmSplash.pgbEntrada.Value / frmSplash.pgbEntrada.MaxProgress) * 100
    'frmSplash.pgbEntrada.Text = "Carregando " & Porcentagem & "%"
sair1:
    Exit Sub
err1:
    Resume Next
End Sub




Function GetIni(section, key, arq)
    'section = É o que está entre []
    'key = É o nome que se encontra antes do sinal de igual (=)
    'arq = É o nome do arquivo INI
    
    Dim val As String
    Dim valor As Integer
    
    val = String$(255, 0)
    valor = GetPrivateProfileString(section, key, "", val, Len(val), arq)
    
    If valor = 0 Then
    GetIni = ""
    Else
    GetIni = Left(val, valor)
    End If

End Function


Sub txt2list(arq, ByRef list As Object)

    Dim s As String
    Open (App.Path & "\" & arq) For Input As #1
    Do Until EOF(1)
    Line Input #1, s
        list.AddItem s
    Loop
    Close #1

End Sub


Sub list2txt(ByRef list As Object, arq)
    Open (App.Path & "\" & arq) For Output As #1
        Dim I As Integer
        With list
            For I = 0 To .ListCount - 1
                Print #1, UCase(.list(I))
            Next
        End With
    Close #1
End Sub




'Cria o Relatorio Tripa p/ as Lojas
Public Sub Criar_RPT_TRIPA(ByRef ado, ByRef AdoItem)
    Dim xlA As New Excel.Application
 '   Dim xlW As New Excel.Workbook
 '   Dim xlP As New Excel.Worksheet
    
    Dim wFunc As Integer
    Dim strLogo As String

On Error GoTo err1
    
 
    'Cria Aplication
    Set xlA = CreateObject("Excel.application")
    'Abrir o arquivo do Excel
  '  Set xlW = xl.Workbooks.Add(Template:=strDirRPT & "\tripa.xlt")
    Set xL = xlA.Workbooks.Add(Template:=strDirRPT & "\tripa.xlt")
    
    
    
    'envia dados das caixa de textos Dev1,....
    'para as celulas coordenadas Ex: Cells(linha, coluna)
'        xlP1.Worksheets("Loja").Cells(5, 2).Value = "teste"
    
w_P = 1
    
    
'Looping entre os Funcionáios
Do While Not ado.EOF
        
    wFunc = wFunc + 1
    
    If Mid(strLogo, 1, 2) <> ado.Fields("logo") Then
        wFunc = 1
        w_P = 1
    End If
    
    If wFunc > 9 Then
      w_P = w_P + 1
      wFunc = 1
      strLogo = ado.Fields("logo") & "_" & w_P
      CopyPlan xL, strLogo, wFunc
    End If
    
    
    
    'Inseri os dados até 9 funcionarios
'    If wFunc <= 9 Then
    
        ' definir qual a planilha de trabalho
        If ado.RecordCount > 1 Then
            'TENTA ABRIR A PLANILHA    SE NÃO EXISTIR CRIA
            If w_P = 1 Then
                strLogo = ado.Fields("logo")
            End If
            
            For p = 1 To xL.Sheets.Count
                
                If xL.Sheets.Item(p).Name = strLogo Then
                    Exit For
                End If
            Next p
            If xL.Sheets.Item(xL.Sheets.Count).Name = strLogo Then
                xL.Sheets(xL.Sheets.Count).Select
            Else
                CopyPlan xL, strLogo, wFunc
            
                'Colocar o Total de Vendas
                xlA.Goto Reference:="TV"
                de.rscmdSqlTotalVND.Filter = "Logo = '" & ado.Fields("logo") & "'"
                If de.rscmdSqlTotalVND.RecordCount > 0 Then xL.Sheets.Application.ActiveCell.value = de.rscmdSqlTotalVND.Fields("TOTAL")
                
            End If
            
            xL.Sheets(strLogo).Select
        Else
            xL.Sheets("Individual").Select
        End If
    
    
    
            'Coloca o Nome e LOGO
            xlA.Goto Reference:="F." & wFunc
            On Error Resume Next
            w_Nome = de.cncDBase.Execute("SELECT NOME FROM LOJB011 WHERE COD_FUNC = '" & ado.Fields("COD_CENTRAL") & "' AND COD_LOJ = '" & ado.Fields("Logo") & "'").Fields("NOME")
            If w_Nome = Empty Then w_Nome = ado.Fields("NOME")
        On Error GoTo err1
            xL.Sheets.Application.ActiveCell.value = ado.Fields("Logo") & " - " & Mid(w_Nome, 1, 20)
        
        
            'Inserir as linhas
            If AdoItem.RecordCount > 15 Then
                    MsgBox "Os itens do relatorio estão sobrecarregados!" & Chr(13) & "Emp : " & ado.Fields("Logo") & " - " & ado.Fields("Nome")
            Else
                
                W_QT = AdoItem.RecordCount
                'Fitlro das Três Primeiras opões na planilha
                AdoItem.Filter = "C_TP_CONTA = 20 OR C_TP_CONTA = 21 OR C_TP_CONTA = 24"
                Do While Not AdoItem.EOF
                        'COMISSÃO
                        Select Case AdoItem.Fields("C_TP_CONTA")
                        Case 20:
                            xlA.Goto Reference:="F." & wFunc & "." & 1
                            xL.Sheets.Application.ActiveCell.FormulaR1C1 = Mid(AdoItem.Fields("Conta"), 1, 15)
                            xlA.Goto Reference:="F." & wFunc & "." & 1 & ".v"
                            xL.Sheets.Application.ActiveCell.FormulaR1C1 = AdoItem.Fields("Valor")
                        Case 21:
                            'PREMIO
                                xlA.Goto Reference:="F." & wFunc & "." & 2
                                xL.Sheets.Application.ActiveCell.FormulaR1C1 = Mid(AdoItem.Fields("Conta"), 1, 15)
                                xlA.Goto Reference:="F." & wFunc & "." & 2 & ".v"
                                xL.Sheets.Application.ActiveCell.FormulaR1C1 = AdoItem.Fields("Valor")
                        Case 24:
                            'Férias
                                xlA.Goto Reference:="F." & wFunc & "." & 3
                                xL.Sheets.Application.ActiveCell.FormulaR1C1 = Mid(AdoItem.Fields("Conta"), 1, 15)
                                xlA.Goto Reference:="F." & wFunc & "." & 3 & ".v"
                                xL.Sheets.Application.ActiveCell.FormulaR1C1 = AdoItem.Fields("Valor")
                        End Select
                
                    AdoItem.MoveNext
                Loop
                
                
                'TIRA O FITLRO
                AdoItem.Filter = "C_TP_CONTA <> 20 AND C_TP_CONTA <> 78 AND C_TP_CONTA <> 21 and C_TP_CONTA <> 24 and C_TP_CONTA <> 31"
                If AdoItem.RecordCount > 0 Then
                    For Linha = 4 To W_QT + 3
                                            
                            'Descrição
                            xlA.Goto Reference:="F." & wFunc & "." & Linha
                            xL.Sheets.Application.ActiveCell.FormulaR1C1 = Mid(AdoItem.Fields("Conta"), 1, 15)
                            
                            'Valor
                            xlA.Goto Reference:="F." & wFunc & "." & Linha & ".v"
                            '*** SE TIPO É   "="  LANÇA COMO TEXTO,  SENÃO COMO NUMERO
                            If AdoItem.Fields("OP") = "=" Then
                                xL.Sheets.Application.ActiveCell.FormulaR1C1 = "'" & AdoItem.Fields("Valor")
                            Else
                                xL.Sheets.Application.ActiveCell.FormulaR1C1 = AdoItem.Fields("Valor")
                            End If
                            
                            If xL.Sheets.Application.ActiveCell.FormulaR1C1 < 0 Then xlA.Selection.Font.ColorIndex = 3
                        
                        AdoItem.MoveNext
                        If AdoItem.EOF Then Exit For
                    Next Linha
                End If
            End If
'    Else
'        MsgBox "O limite de 9 Emp. foi excedido!", vbExclamation
'    End If
    
    
    'Se total Negativo muda a cor p/ Vermelho
    xlA.Goto Reference:="F." & wFunc & ".TOTAL"
    If xL.Sheets.Application.ActiveCell.text < 0 Then xL.Sheets.Application.Selection.Font.ColorIndex = 3

    'Se for maior q/ 9   gera outra planilha,    senão vai p/ prox. reg.
    If wFunc <= 9 And w_P = 1 Then
      w_P = 1
    Else
      'w_P = 2
    End If
    
    
 ado.MoveNext
Loop
  
'DESATIVA O DISPLAY DE ALERTA DO EXCEL
xlA.DisplayAlerts = False
    
    If xL.Sheets.Count = 2 Then
         'Exlui plan base
         xL.Sheets("LOJA").Select
         xlA.ActiveWindow.SelectedSheets.Delete
    Else
         'Exlui plan base
         xL.Sheets("Individual").Select
         xlA.ActiveWindow.SelectedSheets.Delete
         'Exlui plan base
         xL.Sheets("LOJA").Select
         xlA.ActiveWindow.SelectedSheets.Delete
    End If
    
xlA.DisplayAlerts = True
    
    
    xlA.Visible = True

sair:
    Exit Sub
err1:
    xlA.Visible = True
    
    
    MsgBox Err.Number & " : " & Err.Description
    Resume sair
End Sub

Public Sub CopyPlan(ByRef w_xl, ByRef w_strlogo, ByRef wFunc)
On erro GoTo err1
    
    'CRIA UMA NOVA PLANILHA P/ A LOJA
    wFunc = 1
    w_xl.Sheets("LOJA").Select
    w_xl.Sheets("LOJA").Copy After:=w_xl.Sheets(w_xl.Sheets.Count)
    w_xl.Sheets(w_xl.Sheets.Count).Select
    w_xl.Sheets(w_xl.Sheets.Count).Name = w_strlogo

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description
    Resume sair
End Sub



Public Function CALC_PG_EMP(ByRef W_ADO_EMP, ByRef W_DT_PG) As Currency
 
 W_JURO_AO_MES = 0
 W_JURO_AO_DIA = 0
 W_DT_ULT_PG = "00:00:00"
 W_PARC_RESTANTE = 0
 W_SALDO = 0
 w_qt_dias = 0
 W_QT_MESES = 0
 W_SALDO_AC = 0


    W_JURO_AO_MES = IIf(IsNull(W_ADO_EMP.Fields("E_JURO_AO_MES")), 0, W_ADO_EMP.Fields("E_JURO_AO_MES"))
    W_JURO_AO_DIA = IIf(IsNull(W_ADO_EMP.Fields("E_JURO_AO_DIA")), 0, W_ADO_EMP.Fields("E_JURO_AO_DIA"))
    
    If W_JURO_AO_MES <> "0" Then W_JURO_AO_MES = W_JURO_AO_MES / 100
    If W_JURO_AO_DIA <> "0" Then W_JURO_AO_DIA = W_JURO_AO_DIA / 100
    
    If W_JURO_AO_DIA = "0" And W_JURO_AO_MES > 0 Then
        W_JURO_AO_DIA = W_ADO_EMP.Fields("E_JURO_AO_MES") / 30
        de.cnc.Execute "UPDATE TAB_EMPRESTIMO SET E_JURO_AO_DIA = '" & CDbl(W_JURO_AO_DIA) & "' WHERE (E_CODIGO = " & W_ADO_EMP.Fields("E_CODIGO") & ")"
    End If
    
    W_DT_ULT_PG = W_ADO_EMP.Fields("E_DT_ULT_PG")
    W_PARC_RESTANTE = W_ADO_EMP.Fields("E_QT_PARC") - W_ADO_EMP.Fields("E_QT_PG")
    W_SALDO = W_ADO_EMP.Fields("E_SALDO")
    w_qt_dias = CDbl(CVDate(W_DT_PG) - CVDate(W_DT_ULT_PG))
    W_QT_MESES = CDbl(w_qt_dias) / 30
    
    W_SALDO_AC = W_SALDO

    '*** SE O TEMPO P/ CALCULO FOR DE MESES
    If Int(W_QT_MESES) >= 1 Then
        '*** CALCULE JURO SOBRE JURO AO MES
        For I = 1 To Int(W_QT_MESES)  ' INT(W_QT_MES)  SIGNIFICA QTDE DE MESES P / CALCULAR
           W_SALDO_AC = (W_SALDO_AC * W_JURO_AO_MES) + W_SALDO_AC
        Next I
    'End If

    '*** SE SOBRARAM DIAS P/ CALCULO DE JUROS CALCULE
    ElseIf (W_QT_MESES - Int(W_QT_MESES)) > 0 Then
       W_SALDO_AC = (W_SALDO_AC * (W_JURO_AO_DIA * ((W_QT_MESES - Int(W_QT_MESES)) * 30))) + W_SALDO_AC
    End If
    
    CALC_PG_EMP = W_SALDO_AC - W_SALDO
    
End Function

Public Function UltDiaMes(wMes As Byte, wAno As Integer) As Date
Dim w_DtI, w_DtF As Date

    w_DtI = CVDate("01/" & Format(wMes, "00") & "/" & Format(wAno, "0000"))
    w_DtF = w_DtI + 32
    w_DtF = CVDate("01/" & Format(w_DtF, "mm/yyyy")) - 1
    UltDiaMes = w_DtF
    
End Function


'Procedimento que Gera o Arquivo da Tripinha
Public Sub PrintTripa(ByRef ado, ByRef AdoItem, Individual As Boolean)
Dim w_ado_Lojb011 As ADODB.Recordset
Dim w_adoLogo As ADODB.Recordset
Dim fs, a As Object
Dim n As Byte
Dim W_TCPP, W_TC, W_TP, W_TPs As Double
Dim wMaxNivel As Byte
Dim w_arrClass
Dim w_logo As String
Dim w_TF_CLASS As Boolean
Dim W_QT_CX_VND As Byte
Dim w_Saldo_S As Double
On Error GoTo err1
    
Const wCols = 47                    'tamanho em colunas da tripa
Const wColsV = 10                   'Espaço de caracteres reservado p/ o valor
Const wColsT = wCols - wColsV       'Espaço de caracteres reservado p/ o Texto da conta
    
    'Cria Arquivo texto para Impressão da Tripa
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.createTextFile(strDirRPT & "\Tripa.txt")

    Set w_ado_Lojb011 = de.cncDBase.Execute("SELECT NOME, COD_FUNC, COD_LOJ  FROM LOJB011").Clone
    Set w_adoLogo = ado.Clone
    
    wMaxNivel = de.cnc.Execute("SELECT MAX(TP_NIVEL)  FROM TAB_TP_CONTA").Fields(0)
    
    'a.Writeline (String(1, Chr(15)))  'Deixa a letra pequena
    
    '         COLUNA (nOME, vALOR) ,   LINHAS
    
    'Inseri os Dados Na tripa
    Do While Not ado.EOF
        If w_logo <> ado.Fields("Logo") Then
            w_logo = ado.Fields("Logo")
            
            'a.Writeline (ado.Fields("Logo") & " " & String(18, "*") & " " & ado.Fields("Logo") & " " & String(19, "*") & " " & ado.Fields("Logo"))
            'a.Writeline ("")
            'a.Writeline ("")
            
'            w_adoLogo.Filter = "Logo = '" & w_logo & "'"
            ReDim w_arrClass(5, 0)
            W_POS = 0
            W_QT_CX_VND = 0
        End If
                        
        'PEGA A QTDE DE VND P/ FAZER A MÉDIA
        If ado.Fields("TIPO") = "C" Then W_QT_CX_VND = IIf(IsNull(ado.Fields("CX_QT_VND")) Or ado.Fields("CX_QT_VND") = 0, 3, ado.Fields("CX_QT_VND"))
                
        ReDim Preserve w_arrClass(5, W_POS)
        W_POS = W_POS + 1
        
        'Pega o Nome da Vendedora na Central
        If ado.Fields("COD_CENTRAL") <> "" And ado.Fields("Logo") <> "" Then
            w_ado_Lojb011.Filter = "COD_FUNC = '" & ado.Fields("COD_CENTRAL") & "' AND COD_LOJ = '" & ado.Fields("Logo") & "'"
            w_Nome = Left(IIf(w_ado_Lojb011.RecordCount = 0, ado.Fields("Nome"), w_ado_Lojb011.Fields("NOME")), 35)
        Else
            w_Nome = Left(ado.Fields("Nome"), 35)
        End If
    
        If ado.Fields("TIPO") = "C" Then w_Nome = "Extra :  " & w_Nome
    
        a.Writeline (String(wCols, "-"))
        'If ((wCols - Len(w_Nome) - 7) < 0) Then
            'a.Writeline ("| " & ado.Fields("Logo") & " - " & w_Nome & " " & String(wCols - Len(ado.Fields("Logo") & w_Nome), " ") & "|")
            a.Writeline ("| " & w_Nome & " " & String(wCols - Len("| " & w_Nome & " "), " ") & "|")
        'Else
           'a.Writeline ("| " & ado.Fields("Logo") & " - " & w_Nome & " " & String(wCols - Len(ado.Fields("Logo") & w_Nome) - 7, " ") & "|")
        '   a.Writeline ("| " & w_Nome & " " & String(wCols - Len("| " & w_Nome & " ") - 7, " ") & "|")
        'End If
        a.Writeline (String(wCols, "-"))
        
        W_TOTAL = 0
        W_TOTAL_NIVEL = 0
        W_TCPP = 0
        W_TC = 0
        W_TP = 0
        W_TPs = 0
        For n = 0 To wMaxNivel
            AdoItem.Filter = "TP_NIVEL = " & n & ""
            AdoItem.Filter = AdoItem.Filter & " AND FICHA = " & ado.Fields("ficha")
            W_TOTAL_NIVEL = 0
            
            Do While Not AdoItem.EOF
               If AdoItem.Fields("c_tp_conta") <> 111 Then
                   'PEGA A QTDE DE ESPAÇOS A ADICIONAR P/ O VALOR FICAR ALINHADO A DIREITA
                   a.Writeline (Left(AdoItem.Fields("c_tp_conta"), wColsT) _
                               & String(wColsT - Len(Left(AdoItem.Fields("c_tp_conta"), wColsT)), " ") _
                               & String(wColsV - Len(Format(AdoItem.Fields("VALOR"), "0.00")), " ") & Format(AdoItem.Fields("VALOR"), "0.00"))
                   
                   'somatoria das TC or  TP or  Piso
                   If AdoItem.Fields("c_tp_conta") = 20 Or AdoItem.Fields("c_tp_conta") = 21 Or AdoItem.Fields("c_tp_conta") = 23 Or AdoItem.Fields("c_tp_conta") = 22 Or AdoItem.Fields("c_tp_conta") = 35 Or AdoItem.Fields("c_tp_conta") = 34 Then W_TCPP = W_TCPP + CDbl(AdoItem.Fields("VALOR"))
                   If AdoItem.Fields("c_tp_conta") = 20 Then W_TC = W_TC + CDbl(AdoItem.Fields("VALOR"))
                   If AdoItem.Fields("c_tp_conta") = 21 Then W_TP = W_TP + CDbl(AdoItem.Fields("VALOR"))
                   If AdoItem.Fields("c_tp_conta") = 23 Then W_TPs = W_TPs + CDbl(AdoItem.Fields("VALOR"))
                
                   'somatoria
                   'If n > 0 And AdoItem.Fields("op") <> "=" Then W_TOTAL = W_TOTAL + CDbl(AdoItem.Fields("VALOR"))
                   If AdoItem.Fields("op") <> "=" Then W_TOTAL = W_TOTAL + CDbl(AdoItem.Fields("VALOR"))
                   If AdoItem.Fields("op") <> "=" Then W_TOTAL_NIVEL = W_TOTAL_NIVEL + CDbl(AdoItem.Fields("VALOR"))
               End If
               AdoItem.MoveNext
            Loop
            
            If W_TOTAL_NIVEL <> 0 And n >= 1 And n <= 3 Then  'TOTAL ACUMULADO
                a.Writeline (String(wCols, "-"))
                a.Writeline ("TOTAL " & n & String(wColsT - Len("TOTAL " & n), " ") _
                          & String(wColsV - Len(Format(W_TOTAL, "0.00")), " ") _
                          & Format(W_TOTAL, "0.00"))
                          
                a.Writeline (String(wCols, "-"))
            ElseIf W_TOTAL_NIVEL <> 0 Then                    'TOTAL DO GRUPO
                a.Writeline (String(wCols, "-"))
                a.Writeline ("TOTAL " & n & String(wColsT - Len("TOTAL " & n), " ") _
                          & String(wColsV - Len(Format(W_TOTAL_NIVEL, "0.00")), " ") _
                          & Format(W_TOTAL_NIVEL, "0.00"))
                          
                a.Writeline (String(wCols, "-"))
            End If
        Next n
                
        'INSERIR Valores( TCPP, W_TOTAL, W_TC, W_TP, w-TPs)  NO ARRAY
        w_arrClass(0, W_POS - 1) = Left(w_Nome, wColsT - 5)  'NOME
        w_arrClass(1, W_POS - 1) = W_TCPP     'VALOR de Tc e Tp  do Vendedor
        w_arrClass(2, W_POS - 1) = W_TOTAL    'Saldo geral do Vendedor
        w_arrClass(3, W_POS - 1) = W_TC       'Valor Comissão Vendedor
        w_arrClass(4, W_POS - 1) = W_TP       'Valor Premio do Vendedor
        w_arrClass(5, W_POS - 1) = W_TPs      'Valor Piso do Vendedor
               
                
        'TOTALIZAÇÃO DO FUNCIONARIO
        a.Writeline ("F : " & String(wColsT - Len("F : "), " ") & String(wColsV - Len(Format(W_TOTAL, "0.00")), " ") & Format(W_TOTAL, "0.00"))
        a.Writeline (String(wCols, "-") & vbCrLf & vbCrLf & vbCrLf)
        
        ado.MoveNext
        
        If Individual = False Then 'Só faz classificação se não For ficha Individual
        
                If ado.EOF Then
                    w_TF_CLASS = True
                Else
                    If w_logo <> ado.Fields("Logo") Then
                        w_TF_CLASS = True
                    Else
                        w_TF_CLASS = False
                    End If
                End If
                
                
                'Pega o array original
                w_arrClass_Orig = w_arrClass
                
                'acabou a loja  -  totaliza
                    If w_TF_CLASS = True Then
                        w_arrClass = CLASSIFICA_ARRAY(w_arrClass, 1, "D")
                        W_TG = 0
                        W_TC = 0
                        W_TP = 0
                        W_TPs = 0
                        'a.Writeline (String(wCols, "-"))
                        'a.Writeline ("| " & w_logo & " - Classificacao" & String(wCols - Len("| " & w_logo & " - Classificação") - 1, " ") & "|")
                        'Lista a Classificação
                        w_Str_CX = Empty
                        w_Total_S = 0
                        w_qtCX = 0
                        For I = 0 To UBound(w_arrClass, 2)
                           If Not UCase(Left(w_arrClass(0, I), 5)) = "CAIXA" Then
                        '        a.Writeline (String(wCols, "-"))
                        '        a.Writeline (Format(I + 1, "00") & " - " & w_arrClass(0, I) _
                        '                   & String(wColsT - Len(Format(I + 1, "00") & " - " & w_arrClass(0, I)), " ") _
                        '                   & String(wColsV - Len(Format(w_arrClass(1, I), "0.00")), " ") _
                        '                   & Format(w_arrClass(1, I), "0.00"))
                                W_TG = W_TG + CDbl(w_arrClass(1, I))
                                W_TC = W_TC + CDbl(w_arrClass(3, I))
                                W_TP = W_TP + CDbl(w_arrClass(4, I))
                                W_TPs = W_TPs + CDbl(w_arrClass(5, I))
                                If w_qtCX < W_QT_CX_VND Then 'i >= 0 And i <= (W_QT_CX_VND - 1) Then
                                    w_qtCX = w_qtCX + 1
                                    w_Str_CX = w_Str_CX & IIf(Len(w_Str_CX) > 0, "+", "") & Format(CDbl(w_arrClass(1, I)), "0.00")
                                    w_Total_S = w_Total_S + CDbl(CDbl(w_arrClass(1, I)))
                                End If
                            End If
                        Next I
                        'a.Writeline (String(wCols, "-"))
                        'a.Writeline ("Total :" & String(wColsT - Len("Total :"), " ") _
                        '            & String(wColsV - Len(Format(W_TG, "0.00")), " ") & Format(W_TG, "0.00"))
                        'a.Writeline (String(wCols, "-"))
                        W_QT_CX_VND = IIf(W_QT_CX_VND = 0, 3, W_QT_CX_VND)
                        If W_QT_CX_VND > 3 Then  'quebra linha
                            w_Str_CX = "(" & w_Str_CX & ")/" & W_QT_CX_VND
                        '    a.Writeline ("S.Cx :")
                        '    a.Writeline (Left(w_Str_CX, wColsT) & String(wColsT - Len(w_Str_CX), " ") _
                        '                & String(wColsV - Len(Format(w_Total_S / W_QT_CX_VND, "0.00")), " ") _
                        '                & Format(w_Total_S / W_QT_CX_VND, "0.00"))
                        Else
                            w_Str_CX = "S.Cx : (" & w_Str_CX & ")/" & W_QT_CX_VND
                        '    a.Writeline (Left(w_Str_CX, wColsT) & String(wColsT - Len(w_Str_CX), " ") _
                        '                & String(wColsV - Len(Format(w_Total_S / W_QT_CX_VND, "0.00")), " ") _
                        '                & Format(w_Total_S / W_QT_CX_VND, "0.00"))
                        End If
                        'a.Writeline (String(wCols, "-"))
                        'a.Writeline (vbCrLf & vbCrLf)
                        
                        'Pega o total de vendas
                        de.rscmdSqlTotalVND.Filter = "Logo = '" & w_logo & "'"
                        w_TV = 0
                        If de.rscmdSqlTotalVND.RecordCount > 0 Then w_TV = de.rscmdSqlTotalVND.Fields("TOTAL")
                        'a.Writeline (String(wCols, "-"))
                        'a.Writeline ("TC :" & String(wColsT - Len("TC :"), " ") _
                        '            & String(wColsV - Len(Format(W_TC, "0.00")), " ") & Format(W_TC, "0.00"))
                        'a.Writeline (String(wCols, "-"))
                        w_Premio = 0
                        w_Premio = de.cnc.Execute("SELECT SUM(PREMIO1+PREMIO2) AS TOTAL FROM LOJB135 WHERE LOJA = '" & w_logo & "'").Fields(0)
                        'a.Writeline ("TP :[" & Format(w_Premio, "0.00") & "]" & String(wColsT - Len("TP :[" & Format(w_Premio, "0.00") & "]"), " ") _
                        '            & String(wColsV - Len(Format(W_TP, "0.00")), " ") & Format(W_TP, "0.00"))
                        'a.Writeline (String(wCols, "-"))
                        'a.Writeline ("TPs:" & String(wColsT - Len("TPs:"), " ") _
                        '            & String(wColsV - Len(Format(W_TPs, "0.00")), " ") & Format(W_TPs, "0.00"))
                        'a.Writeline (String(wCols, "-"))
                        'a.Writeline ("V  :" & String(wColsT - Len("V  :"), " ") _
                        '            & String(wColsV - Len(Format(w_TV, "0.00")), " ") & Format(w_TV, "0.00"))
                        'a.Writeline (String(wCols, "-"))
                        'a.Writeline (vbCrLf & vbCrLf)
                        
                        
                        'Lista Relação de Saldo e Vendedores
                        w_arrClass = w_arrClass_Orig  'CLASSIFICA_ARRAY(w_arrClass, 0, "C")
                        'a.Writeline (String(wCols, "-"))
                        'a.Writeline ("| " & w_logo & " - Desp. EXTRA" & String(wCols - Len("| " & w_logo & " - Desp. EXTRA") - 1, " ") & "|")
                        W_TG = 0
                        For I = 0 To UBound(w_arrClass, 2)
                        '    a.Writeline (String(wCols, "-"))
                        '    a.Writeline (w_arrClass(0, I) _
                        '               & String(wColsT - Len(w_arrClass(0, I)), " ") _
                        '               & String(wColsV - Len(Format(w_arrClass(2, I), "0.00")), " ") _
                        '               & Format(w_arrClass(2, I), "0.00"))
                            If w_arrClass(2, I) > 0 Then W_TG = W_TG + CDbl(w_arrClass(2, I))
                            
                        Next I
        
        
                        'a.Writeline (String(wCols, "-"))
                        'a.Writeline ("GASTOS ANUAIS :" & String(wColsT - Len("GASTOS ANUAIS :"), " ") _
                        '            & String(wColsV - Len(Format(W_TG, "0.00")), " ") & Format(W_TG, "0.00"))
                        'a.Writeline (String(wCols, "-"))
                        'a.Writeline ("| " & w_logo & " - Desp. EXTRA" & String(wCols - Len("| " & w_logo & " - Desp. EXTRA") - 1, " ") & "|")
                        a.Writeline (w_logo & String(wCols - (Len(w_logo) + Len(Format(W_TG, "##0=000")) + 1), " ") & "=" & Format(W_TG, "##0=000"))
                        
                        a.Writeline (vbCrLf & vbCrLf & vbCrLf)
                        a.Writeline (String(wCols, "#"))
                        a.Writeline (vbCrLf & vbCrLf & vbCrLf)
                        
                        'Repetindo totais (duplicado)
                        a.Writeline (String(wCols, "-") & vbCrLf & vbCrLf & vbCrLf)
                        a.Writeline (w_logo & String(wCols - (Len(w_logo) + Len(Format(W_TG, "##0=000")) + 1), " ") & "=" & Format(W_TG, "##0=000"))
                        a.Writeline (vbCrLf & vbCrLf & vbCrLf)
                        a.Writeline (String(wCols, "#"))
                        a.Writeline (vbCrLf & vbCrLf & vbCrLf)
                        
                        w_TF_CLASS = False
                    End If
        End If
    Loop
        
    a.Writeline (vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf)
        
    a.Close
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description
End Sub

'Função que Classifica o Arrays na coluna do Valor por ordem decrescente
'Criado dia 12/05/06  por Rafael Bianchin
Private Function CLASSIFICA_ARRAY(ByRef w_arrClass, wColSort As Byte, Ordem_C_D As String) As Variant
Dim w_arrOrdem
Dim Ordem_Cresc As Boolean
Dim w_Hab As Boolean

     w_arrOrdem = w_arrClass
     Ordem_Cresc = (Ordem_C_D = "C")
        
    
    For L = 0 To UBound(w_arrClass, 2)
        For LL = L + 1 To UBound(w_arrClass, 2)
            w_Hab = False
            If Ordem_Cresc = True Then 'ordem Crescente
              If w_arrClass(wColSort, L) > w_arrClass(wColSort, LL) Then w_Hab = True
            Else    'ordem Crescente
              If w_arrClass(wColSort, L) < w_arrClass(wColSort, LL) Then w_Hab = True
            End If
            
            If w_Hab = True Then  'habilitado p/  fazer troca
                
                w_0 = w_arrClass(0, L)
                w_1 = IIf(IsEmpty(w_arrClass(1, L)), 0, w_arrClass(1, L))
                w_2 = IIf(IsEmpty(w_arrClass(2, L)), 0, w_arrClass(2, L))
                w_3 = IIf(IsEmpty(w_arrClass(3, L)), 0, w_arrClass(3, L))
                w_4 = IIf(IsEmpty(w_arrClass(4, L)), 0, w_arrClass(4, L))
                w_5 = IIf(IsEmpty(w_arrClass(5, L)), 0, w_arrClass(5, L))
                
                w_arrClass(0, L) = w_arrClass(0, LL)
                w_arrClass(1, L) = IIf(IsEmpty(w_arrClass(1, LL)), 0, w_arrClass(1, LL))
                w_arrClass(2, L) = IIf(IsEmpty(w_arrClass(2, LL)), 0, w_arrClass(2, LL))
                w_arrClass(3, L) = IIf(IsEmpty(w_arrClass(3, LL)), 0, w_arrClass(3, LL))
                w_arrClass(4, L) = IIf(IsEmpty(w_arrClass(4, LL)), 0, w_arrClass(4, LL))
                w_arrClass(5, L) = IIf(IsEmpty(w_arrClass(5, LL)), 0, w_arrClass(5, LL))
                
                w_arrClass(0, LL) = w_0
                w_arrClass(1, LL) = w_1
                w_arrClass(2, LL) = w_2
                w_arrClass(3, LL) = w_3
                w_arrClass(4, LL) = w_4
                w_arrClass(5, LL) = w_5
                
            End If
        Next LL
        
    Next L
    
    CLASSIFICA_ARRAY = w_arrClass
End Function


Public Function ExecuteSQL(SQLString, Optional ByRef w_RegAf, Optional w_Provider As String, Optional w_ShowProgBar As Boolean = True) As ADODB.Recordset
    Dim I As Byte, w_Err As Byte
    
conectar:
On Error GoTo errCNC
    If Not (InStr(UCase(SQLString), "SELECT") > 0 Or InStr(UCase(SQLString), "CALL")) Then w_ShowProgBar = False
    
    
   'Se o tempo da ultima conexao p/ a q/ será realizada for maior q/ o timeOut
   'entao -  feche e abra a conexao
    If (w_Provider <> "" And w_Provider <> i_Provider) Or (timeOut <= Format(Time() - timeCNC, "hh:mm:ss") And timeCNC <> "00:00:00") Then
        Call FechaConexao(Conexão)
        Call AbreConexao(Conexão, w_Provider)
    End If
    
   timeCNC = Time()
    
    '2º Executando SQL
    w_RegAf = 0
    If InStr(UCase(SQLString), "SELECT") > 0 Or InStr(UCase(SQLString), "CALL") Then
        Set ExecuteSQL = Conexão.Execute(SQLString, w_RegAf).Clone
        w_RegAf = ExecuteSQL.RecordCount
        Set ExecuteSQL.ActiveConnection = Nothing
    Else
        If InStr(SQLString, "DROP") Then On Error Resume Next
        'On Error Resume Next
        Conexão.Execute SQLString, w_RegAf
    End If
sair:
    Exit Function
errCNC:

        Call FechaConexao(Conexão)
        Call AbreConexao(Conexão, w_Provider)
    timeCNC = Time
    w_Err = w_Err + 1
    Pause 0.3
    If w_Err <= 5 Then Resume conectar
    MsgBox (Err), vbCritical ', "ExecuteSQL"
    'MsgBox "Erro no ExecutarSQL " & Chr(13) & Chr(13) & "Err: " & Error$, vbCritical, "ExecuteSQL"
    Resume sair
End Function

Private Sub FechaConexao(ByRef Conexão)
On Error Resume Next
        Conexão.Close
        Set Conexão = Nothing
End Sub
Public Sub AbreConexao(ByRef Conexão, Optional ByRef w_Provider)
On Error Resume Next
    '1º - Abrindo Conexão
    Set Conexão = New ADODB.Connection
    
    If (Not w_Provider = "" And i_Provider <> w_Provider) Or w_Provider = "MSDataShape" Then
        Conexão.Provider = w_Provider
    Else
        Conexão.Provider = "MSDASQL.1"
    End If
    i_Provider = w_Provider
    
    Conexão.CursorLocation = adUseClient
    Conexão.Open strConectaFirebird
End Sub

Public Sub Put_System_CNC()

    strConectaFirebird = "Provider=MSDASQL.1;Persist Security Info=true;Data Source=mwts"

End Sub

   'strConectaMySQL = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & strBDHost & _
   '                   ";PORT=3306" & _
   '                   ";DATABASE=" & strBDDataBase & _
   '                   "; USER=" & strBDUser & _
   '                   ";PASSWORD=" & strBDPW & _
   '                   ";OPTION=3;"


Public Function Backup()
Dim w_Access As Access.Application
Set w_Access = New Access.Application

On Error Resume Next
Inicio:

    
    Dim w_Destino As String
    
    w_Destino = InputBox("Entre com o destino onde deseja salvar o backup :", "Backup", "F:\")
    If w_Destino <> "" Then
        CopiarArq strDirBase, "C:\BackupG.mdb"
        
        w_Access.OpenCurrentDatabase "C:\BackupG.mdb", True
        w_Access.CloseCurrentDatabase '*** Fecha a Conexão com o Banco
        w_Access.DBEngine.CompactDatabase "C:\BackupG.mdb", "C:\BackupGC.mdb"
        CopiarArq "C:\BackupG.mdb", w_Destino
        MsgBox "Backup feito com sucesso!", vbInformation
    Else
        MsgBox "Não foi possível fazer o backup!", vbCritical
    End If
    
sair:
    
    Exit Function
err1:
    If Err.Number = 3044 Then
        MsgBox "O SERVIDOR DA CENTRAL DE ONDE SERÃO IMPORTADAS AS TABELAS, DEVE ESTAR DESLIGADO OU O CAMINHO ESTA INCORRETO!", vbCritical
    Else
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    End If
    Resume sair
End Function


Function acessoTotal() As Boolean
    If (UCase(w_usuario) = UCase(NomeMestre) Or UCase(w_usuario) = UCase(NomeMestre2) Or UCase(w_usuario) = UCase(NomeMestre3)) Or UCase(w_usuario) = "BEL" Or UCase(w_usuario) = "KELEN" Or UCase(w_usuario) = "RODRIGO" Then
        acessoTotal = True
    Else
        acessoTotal = False
    End If
End Function

Function calculacpf(CPF As String) As Boolean
    'Esta rotina foi adaptada da revista Fórum Access
    On Error GoTo Err_CPF
    Dim I As Integer 'utilizada nos FOR... NEXT
    Dim strcampo As String 'armazena do CPF que será utilizada para o cálculo
    Dim strCaracter As String 'armazena os digitos do CPF da direita para a esquerda
    Dim intNumero As Integer 'armazena o digito separado para cálculo (uma a um)
    Dim intMais As Integer 'armazena o digito específico multiplicado pela sua base
    Dim lngSoma As Long 'armazena a soma dos digitos multiplicados pela sua base(intmais)
    Dim dblDivisao As Double 'armazena a divisão dos digitos*base por 11
    Dim lngInteiro As Long 'armazena inteiro da divisão
    Dim intResto As Integer 'armazena o resto
    Dim intDig1 As Integer 'armazena o 1º digito verificador
    Dim intDig2 As Integer 'armazena o 2º digito verificador
    Dim strConf As String 'armazena o digito verificador
    
    lngSoma = 0
    intNumero = 0
    intMais = 0
    strcampo = Left(CPF, 9)
    
    'Inicia cálculos do 1º dígito
    For I = 2 To 10
        strCaracter = Right(strcampo, I - 1)
        intNumero = Left(strCaracter, 1)
        intMais = intNumero * I
        lngSoma = lngSoma + intMais
    Next I
    dblDivisao = lngSoma / 11
    
    lngInteiro = Int(dblDivisao) * 11
    intResto = lngSoma - lngInteiro
    If intResto = 0 Or intResto = 1 Then
        intDig1 = 0
    Else
        intDig1 = 11 - intResto
    End If
    
    strcampo = strcampo & intDig1 'concatena o CPF com o primeiro digito verificador
    lngSoma = 0
    intNumero = 0
    intMais = 0
    'Inicia cálculos do 2º dígito
    For I = 2 To 11
        strCaracter = Right(strcampo, I - 1)
        intNumero = Left(strCaracter, 1)
        intMais = intNumero * I
        lngSoma = lngSoma + intMais
    Next I
    dblDivisao = lngSoma / 11
    lngInteiro = Int(dblDivisao) * 11
    intResto = lngSoma - lngInteiro
    If intResto = 0 Or intResto = 1 Then
        intDig2 = 0
    Else
        intDig2 = 11 - intResto
    End If
    strConf = intDig1 & intDig2
    'Caso o CPF esteja errado dispara a mensagem
    If strConf <> Right(CPF, 2) Then
        calculacpf = False
    Else
        calculacpf = True
    End If
    Exit Function
    
Exit_CPF:
        Exit Function
Err_CPF:
        MsgBox Error$
        Resume Exit_CPF
End Function

Public Sub Sendkeys(text$, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys text, wait
   Set WshShell = Nothing
End Sub
