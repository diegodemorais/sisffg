VERSION 5.00
Begin VB.Form frm_Alteracao_ADO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALTERAÇÃO DE FICHA MENSAL"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   Icon            =   "frm_Alteracao_ADO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXT_NFICHA 
      Alignment       =   1  'Right Justify
      DataField       =   "M_NFICHA"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
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
      Left            =   360
      TabIndex        =   12
      Top             =   960
      Width           =   1050
   End
   Begin VB.ComboBox TXT_MES 
      DataField       =   "M_MES"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
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
      ItemData        =   "frm_Alteracao_ADO.frx":1CFA
      Left            =   3480
      List            =   "frm_Alteracao_ADO.frx":1D22
      TabIndex        =   6
      Top             =   960
      Width           =   780
   End
   Begin VB.TextBox TXT_FERIAS 
      DataField       =   "M_FERIAS"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
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
      TabIndex        =   4
      Top             =   2520
      Width           =   4695
   End
   Begin VB.TextBox TXT_OBS 
      DataField       =   "M_OBS"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
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
      TabIndex        =   5
      Top             =   3795
      Width           =   4695
   End
   Begin VB.TextBox TXT_ANO 
      DataField       =   "M_ANO"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
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
      Left            =   4245
      TabIndex        =   2
      Top             =   960
      Width           =   810
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H80000009&
      Height          =   5775
      Left            =   5520
      ScaleHeight     =   5715
      ScaleWidth      =   3855
      TabIndex        =   0
      Top             =   600
      Width           =   3915
   End
   Begin VB.PictureBox ADOREG 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   9435
      TabIndex        =   14
      Top             =   6465
      Width           =   9495
   End
   Begin VB.PictureBox ImageList1 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   4440
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   15
      Top             =   0
      Width           =   1200
   End
   Begin VB.PictureBox BarraF 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   9435
      TabIndex        =   1
      Top             =   0
      Width           =   9495
   End
   Begin VB.PictureBox TXT_FUNC 
      DataField       =   "M_F_COD"
      DataSource      =   "ADOREG"
      Enabled         =   0   'False
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
      Left            =   360
      ScaleHeight     =   300
      ScaleWidth      =   4635
      TabIndex        =   3
      Top             =   1800
      Width           =   4695
   End
   Begin VB.PictureBox ADO_GRID 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   9435
      TabIndex        =   16
      Top             =   6135
      Visible         =   0   'False
      Width           =   9495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº FICHA"
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
      TabIndex        =   13
      Top             =   720
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
      Left            =   3600
      TabIndex        =   11
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FÉRIAS"
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
      TabIndex        =   10
      Top             =   2280
      Width           =   2175
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
      TabIndex        =   9
      Top             =   3555
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "FUNCIONÁRIO"
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
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
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
      Left            =   4440
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   4335
      Left            =   120
      Top             =   600
      Width           =   5175
   End
End
Attribute VB_Name = "frm_Alteracao_ADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W_LD_FILTRO As Boolean
Dim V_MOVE As Boolean
Dim V_MOVE_GRID As Boolean


Private Sub Form_Load()
On Error GoTo err1
    
    de.rscmdBase.Close
    de.rscmdBase.Open "SELECT TAB_FICHA_MENS.*  FROM TAB_FICHA_MENS Order By TAB_FICHA_MENS.M_ANO, TAB_FICHA_MENS.M_Mes", , adOpenStatic, adLockOptimistic

    Set ADOREG.Recordset = de.rscmdBase.Clone
    
    Set ADO_GRID.Recordset = de.cnc.Execute("SELECT TAB_FICHA_MENS.* , TAB_FUNCIONARIO.F_NOME FROM TAB_FUNCIONARIO, TAB_FICHA_MENS WHERE TAB_FUNCIONARIO.F_CODIGO = TAB_FICHA_MENS.M_F_COD Order By TAB_FICHA_MENS.M_ANO, TAB_FICHA_MENS.M_Mes").Clone
    V_MOVE = True
    'ADOREG.Refresh
'    de.rsTAB_FICHA_MENS.Close


sair:
    Exit Sub
err1:
    If Not Err.Number = 3705 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

'*** Caption no navegador ***
Private Sub ADOREG_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1
    If Not ADOREG.Recordset.EOF Then ADOREG.Caption = "REGISTRO : " & ADOREG.Recordset.AbsolutePosition & " / " & ADOREG.Recordset.RecordCount & IIf(W_LD_FILTRO = True, " (FILTRADO)", "")
    
   '*** DESABILITA O EDITAR ****
   If ADOREG.Recordset.Fields("M_BLOQ") = True Then
        BarraF.Buttons("editar").Enabled = False
   Else
        BarraF.Buttons("editar").Enabled = True
   End If
    
    
   If V_MOVE = True Then
        On Error Resume Next
        V_MOVE = False
        'ADO_GRID.Recordset.Requery
        If Not ADO_GRID.Recordset.EOF Then
            
            Select Case adReason
            Case 12: '*** Vai p/ o Primeiro Registro ***
                ADO_GRID.Recordset.MoveFirst
            Case 13: '*** Vai p/ o Próximo Registro ***
                ADO_GRID.Recordset.MoveNext
            Case 14: '*** Vai p/ o Anterior Registro ***
                ADO_GRID.Recordset.MovePrevious
            Case 15: '*** Vai p/ o Ultimo Registro ***
                ADO_GRID.Recordset.MoveLast
            
            End Select
                
        End If
   End If
   
sair:
    V_MOVE = True
    Exit Sub
err1:
    If Not Err.Number = -2147217885 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub ADO_GRID_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1

    If V_MOVE = True Then
        V_MOVE = False
        ADOREG.Recordset.Requery
        ADOREG.Recordset.Move ADO_GRID.Recordset.AbsolutePosition - 1
        V_MOVE = True
    End If

sair:
    V_MOVE = True
    Exit Sub
err1:
    If Not (Err.Number = -2147217885) And Not (Err.Number = 3021) And Not (Err.Number = 91) Then MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


'** Barra de Ferramenta ***
Private Sub BarraF_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "fechar": Fechar
        Case "editar": Editar
        Case "salvar": Salvar
        Case "cancelar": Cancelar
        Case "excluir": Excluir
        Case "filtrar": FILTRAR
    End Select
End Sub


'*** Rotinas ***
Private Sub Cancelar()
On Error GoTo err1
    
    POS = ADOREG.Recordset.AbsolutePosition - 1
    ADOREG.Recordset.CancelBatch adAffectCurrent
    ADOREG.Refresh
    ADOREG.Recordset.Move POS

    Editar
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub Editar()
On Error GoTo err1
    
If ADOREG.Recordset.Fields("M_BLOQ") = False Then

    BarraF.Buttons("salvar").Enabled = Not BarraF.Buttons("salvar").Enabled
    BarraF.Buttons("cancelar").Enabled = Not BarraF.Buttons("cancelar").Enabled
    BarraF.Buttons("editar").Enabled = Not BarraF.Buttons("editar").Enabled
    
    Grid.Enabled = Not Grid.Enabled
    
    TXT_MES.Enabled = Not TXT_MES.Enabled
    TXT_ANO.Enabled = Not TXT_ANO.Enabled
    TXT_FUNC.Enabled = Not TXT_FUNC.Enabled
    TXT_FERIAS.Enabled = Not TXT_FERIAS.Enabled
    TXT_OBS.Enabled = Not TXT_OBS.Enabled

    If BarraF.Buttons("salvar").Enabled = False Then
        Grid.SetFocus
    Else
        TXT_MES.SetFocus
    End If

Else
    MsgBox "VOCÊ NÃO PODE ALTERAR FICHA DE MESES PASSADOS!", vbExclamation
End If
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub Excluir()
On Error GoTo err1
        
    If vbYes = MsgBox("DESEJA REALMENTE EXCLUIR A FICHA MENSAL (" & TXT_NFICHA & " : " & TXT_FUNC & ")?", vbQuestion + vbYesNo) Then
        w_pos = ADOREG.Recordset.AbsolutePosition - 1
        ADOREG.Recordset.Delete
        ADOREG.Recordset.UpdateBatch
        w_adoFiltro = ADOREG.Recordset.Filter
        Form_Load
        ADO_GRID.Refresh
        ADOREG.Refresh
      
        ADOREG.Recordset.Filter = w_adoFiltro
        ADO_GRID.Recordset.Filter = w_adoFiltro
      
        Grid.Refresh
      
      '  Grid.ReBind
    End If
    
sair:
    Exit Sub
err1:
    If Not Err.Number = -2147467259 Then
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    Else
        MsgBox "NÃO É POSSÍVEL EXCLUIR ESTA FICHA MENSAL, DEVIDO A CÁLCULOS RELACIONADAS A ELA!", vbCritical
        ADOREG.Refresh
    End If
    Resume sair
End Sub

Private Sub Fechar()
On Error GoTo err1
    Unload Me
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub FILTRAR()
Dim w_resp As String
Dim W_CAMPO As String
Dim W_FILTRO As String
Dim W_FILTRO1 As String

On Error GoTo err1
    
    w_resp = InputBox("FILTRAR PELO QUÊ ? ENTRE COM O NÚMERO DA OPÇÃO DESEJADA." & Chr(13) & Chr(13) & "1 - Nº DA FICHA" & Chr(13) & "2 - FUNCIONÁRIO" & Chr(13) & "3 - MÊS E ANO" & Chr(13) & "4 - REMOVER FILTRO *", , "1")
    
    
    If Not w_resp = "" And IsNumeric(w_resp) And w_resp >= 1 And w_resp <= 4 Then
        Select Case w_resp
        'NFICHA
        Case 1:
            w_resp = "Nº FICHA"
            W_CAMPO = "M_NFICHA"
        'FUNCIONÁRIO
        Case 2:
            w_resp = "FUNCIONÁRIO"
            W_CAMPO = "M_F_COD"
            
        'MÊS E ANO
        Case 3:
            w_resp = "MÊS E ANO"
            W_CAMPO = "M_MES"
            
        '*** REMOVE O FILTRO ****
        Case 4:
            If Not ADOREG.Recordset.Filter = 0 Then
                W_LD_FILTRO = False
                ADOREG.Recordset.Filter = 0
                ADOREG.Refresh
            End If
        End Select
        
        If Not w_resp = "3" Then
            Select Case w_resp
            Case "Nº FICHA":
                W_FILTRO = InputBox("ENTRE COM O " & w_resp & " DESEJADO !")
            Case "FUNCIONÁRIO":
                frm_ESCOLHA_FUNC.Show 1
                W_FILTRO = frm_ESCOLHA_FUNC.TXT_FUNC_COD
            Case "MÊS E ANO":
                W_FILTRO = CDbl(InputBox("ENTRE COM O MÊS DESEJADO !", , Format(Date, "MM")))
                W_FILTRO1 = CDbl(InputBox("ENTRE COM O ANO DESEJADO !", , Format(Date, "YYYY")))
                
                If Not W_FILTRO = "" And IsNumeric(W_FILTRO) And IsNumeric(W_FILTRO1) And Len(W_FILTRO1) = 4 Then
                    W_FILTRO = "M_MES = " & W_FILTRO & " AND M_ANO = " & W_FILTRO1
                    W_LD_FILTRO = True
                    ADOREG.Recordset.Filter = W_FILTRO
                    ADO_GRID.Recordset.MoveFirst
                End If
                                
            End Select
        
                If Not W_FILTRO = "" And IsNumeric(W_FILTRO) Then
                    W_FILTRO = W_CAMPO & " = " & W_FILTRO
                    W_LD_FILTRO = True
                    ADOREG.Recordset.Filter = W_FILTRO
                End If
        
        End If
    End If
    
    ADO_GRID.Refresh
    ADO_GRID.Recordset.Filter = ADOREG.Recordset.Filter
    
    
sair:
    Exit Sub
err1:
    If Err.Number <> 13 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
        W_LD_FILTRO = False
        Resume sair

End Sub

Private Sub Salvar()
On Error GoTo err1
    
    ADOREG.Recordset.UpdateBatch adAffectCurrent
        
    '*** Atualiza o Funcionário ****
    de.cnc.Execute "UPDATE TAB_FUNCIONARIO SET F_FERIAS = '" & TXT_FERIAS & "', F_OBS = '" & TXT_OBS & "' WHERE (F_Codigo = " & TXT_FUNC.BoundText & " )", w_reg
    If w_reg = 0 Then MsgBox "Não foi possível atualiza o cadastro de funcionários (as férias e observações)", vbCritical
    
    Editar
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
    
End Sub










'--------- Ao Pressionar uma Tecla -----------
Private Sub txt_ferias_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub TXT_Obs_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_mes_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub TXT_FUNC_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub TXT_ano_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub GRID_KeyUp(KeyCode As Integer, Shift As Integer)
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
            Editar
    Case 83: ' "S"
            Salvar
    Case 67: ' "C"
            Cancelar
    Case 88: ' "X"
            Excluir
    Case 84: ' "T"
            FILTRAR
    End Select
End If
End Sub


