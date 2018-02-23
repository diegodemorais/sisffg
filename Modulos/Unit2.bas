Attribute VB_Name = "Unit"
Public strDirBase As String  'define o diretório do Banco de dados
Public strImgFundo As String 'define Imagem de fundo
Public strImgSplash As String 'define Imagem de fundo

Declare Function GetPrivateProfileString Lib "Kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As String, ByVal lpDefault As String, ByVal _
lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName _
As String) As Long


Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

'*** CONSTANTES
    '*** MESTRE ***
    Public Const NomeMestre = "MASTER"
    Public Const SenhaMestre = "RPPL"
    
    '*** PROFESSOR ***
    Public Const NomeUsu = "RP"
    Public Const SenhaUsu = "RPFF"

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
        SendKeys "{tab}{home}+{end}"
    Case 38 'seta para baixo
        SendKeys "+{tab}{home}+{end}"
    Case 40 'seta para cima
        SendKeys "{tab}{home}+{end}"
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
    Static w_umaVez As Byte
       If w_umaVez = 0 Then
            'fechar e abrir todos os records set's
            'RODAR NA TELA DE SPLASH - na 1ª vez
                'If frmSplash.pgbEntrada.Value > 0 Then frmSplash.pgbEntrada.Value = 0 'progress bar da tela de splash
               ' frmSplash.pgbEntrada.MaxProgress = de.Recordsets.Count   'progress bar da tela de splash
         
         On Error Resume Next
            For i = 1 To de.Recordsets.Count
                If de.Recordsets(i).State = 1 Then
                   de.Recordsets(i).Close
                End If
                If UCase(Mid(de.Commands.Item(i).Name, 1, 3)) = "TAB" Then de.Recordsets(i).Open
                'ACRESCENTA_BARRA 'ACRESCENTA BARRA DA TELA DE SPLASH
            Next
            
            w_umaVez = 1
       End If
sair:
    Exit Sub
err1:
    MsgBox Error$ & dePP.Recordsets(i).Source, vbCritical
    Resume sair
End Sub






Function GetIni(section, key, arq)
    'section = É o que está entre []
    'key = É o nome que se encontra antes do sinal de igual (=)
    'arq = É o nome do arquivo INI
    
    Dim Val As String
    Dim valor As Integer
    
    Val = String$(255, 0)
    valor = GetPrivateProfileString(section, key, "", Val, Len(Val), arq)
    
    If valor = 0 Then
    GetIni = ""
    Else
    GetIni = Left(Val, valor)
    End If

End Function
