VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_Faltas 
   Caption         =   "FALTAS"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      MaxLength       =   255
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   4080
      TabIndex        =   0
      Top             =   1320
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      MaxSelCount     =   31
      MultiSelect     =   -1  'True
      ShowToday       =   0   'False
      StartOfWeek     =   70385665
      CurrentDate     =   41408
      MaxDate         =   73050
      MinDate         =   2
   End
   Begin MSAdodcLib.Adodc adoReg 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   3750
      Width           =   6630
      _ExtentX        =   11695
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexGRID_L 
      Bindings        =   "frm_Faltas.frx":0000
      Height          =   3615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   10
      FixedRows       =   0
      FixedCols       =   0
      ForeColorSel    =   -2147483639
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
      GridLineWidthFixed=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
   End
End
Attribute VB_Name = "frm_Faltas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wTxtOld As String

'--------- flex grid -------------------------------------
Private ControlVisible As Boolean     ' Se o controle esta visivel ou nao
Private LastRow As Long               ' Ultima linha em que se editou
Private LastCol As Long               ' ultima coluna em que se editou



Private Sub adoReg_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1

If Not IsNumeric(adoReg.Recordset.RecordCount) Then adoReg.Caption = "REGISTRO : " & adoReg.Recordset.AbsolutePosition & " / " & adoReg.Recordset.RecordCount & IIf(W_LD_FILTRO = True, " (FILTRADO)", "")

sair:
    Exit Sub
err1:
    If Not Err.Number = -2147217885 And Not Err.Number = -2147467259 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub Form_Activate()
    flexGRID_L.ColWidth(0) = 0 'codigo
    flexGRID_L.ColWidth(1) = 0 'nficha
    flexGRID_L.ColWidth(2) = 1000 'data (descrição do tipo da conta)
    flexGRID_L.ColWidth(3) = 1000 'atestado
    flexGRID_L.ColWidth(4) = 4000 'obs
End Sub

Private Sub Form_Load()
On Error Resume Next
On Error GoTo err1
   
    If de.rscmdSqlFaltas.State = 1 Then de.rscmdSqlFaltas.Close
    de.cmdSqlFaltas frm_Alt_Fic_Mensal_VIS.TXT_NFICHA
    
sair:
    Set adoReg.Recordset = de.rscmdSqlFaltas.Clone
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair

End Sub






Private Sub flexGRID_L_DblClick()
    'If flexGRID_L.RowSel <> 0 Then CONTA
End Sub




Private Sub flexGRID_L_KeyDown(KeyCode As Integer, Shift As Integer)
    If flexGRID_L.RowSel <> 0 Then
      If (UCase(frmLogin.txtUserName) = UCase(NomeMestre) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre2) Or UCase(frmLogin.txtUserName) = UCase(NomeMestre3)) And Shift = 0 And KeyCode <> 13 Then
          Select Case KeyCode
          End Select
        ElseIf Shift <> 2 And KeyCode = 13 Then

        End If
    End If
End Sub


Private Sub flexGRID_L_KeyPress(KeyAscii As Integer)
    If flexGRID_L.RowSel <> 0 Then
        Select Case KeyAscii
        ' Editar ao teclar ENTER
        Case vbKeyReturn
            KeyAscii = 0
            ExibirCelula
        End Select
    End If
        
End Sub


Private Sub flexGRID_L_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If flexGRID_L.RowSel <> 0 Then
        If Button = 2 Then
            'PopupMenu mnu
        End If
    End If
End Sub





Private Sub ExibirCelula()
    Static OK As Boolean
    
    wTxtOld = ""
    
    ' Se for celula fixa , sair
    If flexGRID_L.Col <= flexGRID_L.FixedCols - 1 Or flexGRID_L.Row <= flexGRID_L.FixedRows - 1 Then
        Exit Sub
    End If
     
    If (flexGRID_L.ColSel <= 1) Or (flexGRID_L.ColSel > 4) Then
        Exit Sub
    End If
    
    If OK Then Exit Sub
    OK = True
    
    OcultarControles
    
    LastRow = flexGRID_L.Row
    LastCol = flexGRID_L.Col
    
    'Nova Celula
    With flexGRID_L
        If .TextMatrix(LastRow, 0) = NovaLinha Then
            .Rows = .Rows + 1
            .TextMatrix(LastRow, 0) = LastRow
            .TextMatrix(.Rows - 1, 0) = NovaLinha
       End If
    End With
    
    Select Case LastCol
    Case Else
        Text1.Move flexGRID_L.CellLeft - Screen.TwipsPerPixelX, flexGRID_L.CellTop - 3 - Screen.TwipsPerPixelY, flexGRID_L.CellWidth + Screen.TwipsPerPixelX * 2, flexGRID_L.CellHeight + Screen.TwipsPerPixelY * 2
        Text1.Text = flexGRID_L.Text
        If Len(flexGRID_L.Text) = 0 Then
            If LastRow > 1 Then
                Text1.Text = flexGRID_L.TextMatrix(LastRow - 1, LastCol)
            End If
        End If
        Text1.Visible = True
        If Text1.Visible Then
            Text1.ZOrder
            Text1.SetFocus
        End If
    End Select
    
    ControlVisible = True
    OK = False
    
    wTxtOld = Text1.Text

End Sub
Private Sub ProximaCelula()
    If flexGRID_L.Col < flexGRID_L.Cols - 1 Then
        flexGRID_L.Col = flexGRID_L.Col + 1
    Else
        flexGRID_L.Col = 1
        If flexGRID_L.Row < flexGRID_L.Rows - 1 Then
            flexGRID_L.Row = flexGRID_L.Row + 1
        End If
    End If
End Sub
Private Sub AtribuiValorCelula()
    Dim texto As String
    Dim Op As String
    texto = Text1.Text
    
    If texto <> flexGRID_L.TextMatrix(flexGRID_L.RowSel, flexGRID_L.ColSel) Then 'Se houve alteração
    
        'Op = flexGRID_L.TextMatrix(flexGRID_L.RowSel, 5) 'op
        
        If flexGRID_L.ColSel = 2 Then 'Se Data (e nao digitou data)
            If Not (IsDate(texto)) Then
                MsgBox "Digite uma data válida ou [ESC] para CANCELAR!", vbCritical, "Data inválida"
                Exit Sub
            End If
        End If
        
        If flexGRID_L.ColSel = 3 Then 'Se Atestado (e não digitou boolean)
            If Not (texto = "Sim" Or texto = "Não") Then
                MsgBox "Digite Sim ou Não!", vbCritical, "Texto inválido"
                Exit Sub
            End If
            
        End If
        
       If flexGRID_L.ColSel = 4 Then 'Se Data (e não digitou data)
            If Len(texto) > 255 Then
                MsgBox "Digite uma observação com menos de 255 caracteres ou [ESC] para CANCELAR!", vbCritical, "Texto muito longo"
                Exit Sub
            End If
        End If
        
            
        'If (MsgBox("Deseja salvar as alterações?", vbYesNo, "Gravar alterações") = vbYes) Then

        flexGRID_L.TextMatrix(LastRow, LastCol) = texto
        flexGRID_L.CellForeColor = vbBlue
        
            If flexGRID_L.ColSel = 2 Then 'Data
                de.cnc.Execute ("UPDATE TAB_FALTA set FT_DATA = '" & CDate(texto) & "' WHERE FT_CODIGO = " & Str(flexGRID_L.TextMatrix(flexGRID_L.RowSel, 0)))
                'de.cmdIncluirLog Date, Time, w_usuario, "EDITAR", "LANÇAMENTOS", "FICHA: " & TXT_NFICHA & " | DATA: " & texto & " | VALOR: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 4) & " | CONTA COD: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 1) & " | CONTA E DESCRICAO: " & texto & " | OP: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 5) & "   >>> DATA ANTERIOR: " & wTxtOld
            ElseIf flexGRID_L.ColSel = 3 Then 'Atestado
                de.cnc.Execute ("UPDATE TAB_FALTA set FT_ATESTADO = '" & texto & "' WHERE FT_CODIGO = " & Str(flexGRID_L.TextMatrix(flexGRID_L.RowSel, 0)))
                'de.cmdIncluirLog Date, Time, w_usuario, "EDITAR", "LANÇAMENTOS", "FICHA: " & TXT_NFICHA & " | DATA: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 0) & " | VALOR: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 4) & " | CONTA COD: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 1) & " | DESCRICAO: " & texto & " | OP: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 5) & "   >>> DESCRIÇÃO ANTERIOR: " & wTxtOld
            ElseIf flexGRID_L.ColSel = 4 Then 'Obs
                de.cnc.Execute ("UPDATE TAB_FALTA set FT_OBS = '" & texto & "' WHERE FT_CODIGO = " & Str(flexGRID_L.TextMatrix(flexGRID_L.RowSel, 0)))
                'de.cmdIncluirLog Date, Time, w_usuario, "EDITAR", "LANÇAMENTOS", "FICHA: " & TXT_NFICHA & " | DATA: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 0) & " | VALOR: " & Str(texto) & " | CONTA COD: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 1) & " | DESCRICAO: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 3) & " | OP: " & flexGRID_L.TextMatrix(flexGRID_L.RowSel, 5) & "   >>> VALOR ANTERIOR: " & wTxtOld
            End If
            
        'End If
    End If
    
    OcultarControles
    ControlVisible = False
    
    'Timer1 = True
    
End Sub
Private Sub OcultarControles()
    ' Ocultar o controle textbox
    Text1.Text = ""
    Text1.Visible = False
End Sub

Private Sub Text1_GotFocus()
    With Text1
         'Seleciona tudo
         .SelStart = 0
         .SelLength = Len(Text1.Text)
         .SetFocus
         
        ' Posiciona o cursor no fim do texto
        '.SelStart = Len(.Text)
    End With
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
    ' ao pressionar ENTER aceitar a entrada de dados
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        AtribuiValorCelula
        'ProximaCelula
    ' ESC, cancela a edição
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Text1.Visible = False
        ControlVisible = False
    End If
End Sub

Private Sub Text1_LostFocus()
    OcultarControles
End Sub
