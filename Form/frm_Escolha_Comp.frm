VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_Escolha_Comp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rel. - Comparativo"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3450
   Icon            =   "frm_Escolha_Comp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Opções : "
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3210
      Begin VB.CheckBox ck 
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
         Left            =   1935
         TabIndex        =   9
         Top             =   1455
         Width           =   975
      End
      Begin VB.OptionButton Op5 
         Caption         =   "@"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Op6 
         Caption         =   "®"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   7
         Top             =   690
         Width           =   855
      End
      Begin VB.OptionButton Op7 
         Caption         =   "(D)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txt_1 
         Height          =   300
         Left            =   360
         TabIndex        =   10
         Top             =   2160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   7
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin VB.OptionButton Op4 
         Caption         =   "Observação"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton Op3 
         Caption         =   "(F)"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton Op2 
         Caption         =   "(B)"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Op1 
         Caption         =   "Anotação"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txt_2 
         Height          =   300
         Left            =   2010
         TabIndex        =   11
         Top             =   2160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   7
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Mes/Ano 2"
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
         Left            =   2010
         TabIndex        =   13
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Mes/Ano 1"
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
         Left            =   360
         TabIndex        =   12
         Top             =   1920
         Width           =   975
      End
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   1482
      ButtonWidth     =   1667
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
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
            Caption         =   "&Imprimir"
            Key             =   "RPT"
            Object.ToolTipText     =   "Imprimir Relatório (Alt+I)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar Alteração (Alt+C)"
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1920
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Escolha_Comp.frx":1CFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Escolha_Comp.frx":2014
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Escolha_Comp.frx":32F6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frm_Escolha_Comp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w_Sql As String


Private Sub BarraF_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err1
   
    Select Case Button.key
        Case "fechar": Fechar
        Case "RPT":
                    If CK.value = 0 Then
                        RPT
                    Else
                        For I = 1 To 7
                            If I = 1 Then
                                Op1_Click
                            ElseIf I = 2 Then
                                Op2_Click
                            ElseIf I = 3 Then
                                Op3_Click
                            ElseIf I = 4 Then
                                Op4_Click
                            ElseIf I = 5 Then
                                Op5_Click
                            ElseIf I = 6 Then
                                Op6_Click
                            ElseIf I = 7 Then
                                Op7_Click
                            End If
                            
                            If I = 1 Then
                                w_Sql2 = w_Sql
                            ElseIf I <= 7 Then
                                w_Sql2 = w_Sql2 & " UNION ALL " & w_Sql
                            End If
                        Next I
                         
                        w_Sql = w_Sql2
                        RPT
                    End If
    End Select

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub




'*** Rotinas ***
Private Sub RPT()
On Error GoTo err1
    
If CK.value = 0 Then
        If Op1.value = True Then
            Op1_Click
            w_Op = "Anotação"
        ElseIf Op2.value = True Then
            Op2_Click
            w_Op = "(B)"
        ElseIf Op3.value = True Then
            Op3_Click
            w_Op = "(F)"
        ElseIf Op4.value = True Then
            Op4_Click
            w_Op = "Observação"
        ElseIf Op5.value = True Then
            Op5_Click
            w_Op = "@"
        ElseIf Op6.value = True Then
            Op6_Click
            w_Op = "®"
        ElseIf Op7.value = True Then
            Op7_Click
            w_Op = "(D)"
        End If
End If
        
        
        If de.rscmdSqlComparativo.State = 1 Then de.rscmdSqlComparativo.Close
        de.rscmdSqlComparativo.Source = w_Sql
        de.rscmdSqlComparativo.Open w_Sql
        de.rscmdSqlComparativo.Sort = "Op, L"

        rptRelComparativo.Sections("SecCab").Controls("LBPER").Caption = "Período :       " & txt_1 & "  à  " & txt_2
        rptRelComparativo.Sections("SecCab").Controls("lbDesc1").Caption = txt_1
        rptRelComparativo.Sections("SecCab").Controls("lbDesc2").Caption = txt_2
        
        If Not de.rscmdSqlComparativo.EOF Then
            rptRelComparativo.Show 1
        Else
            MsgBox w_Op & " - Não existe nenhuma alteração!", vbCritical
        End If
        
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
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



'--------- Ao Pressionar uma Tecla -----------
Private Sub txt_DESC_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_OP_KeyUp(KeyCode As Integer, Shift As Integer)
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

        End Select
    End If
End Sub

Private Sub Form_Load()
    txt_1 = Format(CVDate("01/" & Format(Date, "mm/yyyy")) - 1, "mm/yyyy")
    txt_2 = Format(Date, "mm/yyyy")
End Sub


Private Sub Op1_Click()
    w_Sql = "SELECT 'Anot' as OP, SQL_COMP_1.M_ANOTACAO AS Desc_1, SQL_COMP_2.M_ANOTACAO AS Desc_2, SQL_COMP_1.Mes_Ano AS Per1, SQL_COMP_2.Mes_Ano AS Per2,TAB_FUNCIONARIO.F_Cod_L AS L, TAB_FUNCIONARIO.F_NOME AS N FROM SQL_COMP_1, SQL_COMP_2, TAB_FUNCIONARIO WHERE SQL_COMP_1.M_F_COD = SQL_COMP_2.M_F_COD AND SQL_COMP_2.M_F_COD = TAB_FUNCIONARIO.F_Codigo  AND SQL_COMP_1.M_ANOTACAO <> SQL_COMP_2.M_ANOTACAO AND (SQL_COMP_1.Mes_Ano = '" & txt_1 & "') AND (SQL_COMP_2.Mes_Ano = '" & txt_2 & "')"
End Sub

Private Sub Op2_Click()
    w_Sql = "SELECT '(B)' as OP, SQL_COMP_1.M_LOGO AS Desc_1, SQL_COMP_2.M_LOGO AS Desc_2, SQL_COMP_1.Mes_Ano AS Per1, SQL_COMP_2.Mes_Ano AS Per2, TAB_FUNCIONARIO.F_Cod_L AS L, TAB_FUNCIONARIO.F_NOME AS N FROM SQL_COMP_1, SQL_COMP_2, TAB_FUNCIONARIO WHERE SQL_COMP_1.M_F_COD = SQL_COMP_2.M_F_COD AND SQL_COMP_2.M_F_COD = TAB_FUNCIONARIO.F_Codigo  AND SQL_COMP_1.M_LOGO <> SQL_COMP_2.M_LOGO AND (SQL_COMP_1.Mes_Ano = '" & txt_1 & "') AND (SQL_COMP_2.Mes_Ano = '" & txt_2 & "')"
End Sub

Private Sub Op3_Click()
    w_Sql = "SELECT '(F)' as OP, SQL_COMP_1.M_FERIAS AS Desc_1, SQL_COMP_2.M_FERIAS AS Desc_2, SQL_COMP_1.Mes_Ano AS Per1, SQL_COMP_2.Mes_Ano AS Per2, TAB_FUNCIONARIO.F_Cod_L AS L, TAB_FUNCIONARIO.F_NOME AS N FROM SQL_COMP_1, SQL_COMP_2, TAB_FUNCIONARIO WHERE SQL_COMP_1.M_F_COD = SQL_COMP_2.M_F_COD AND SQL_COMP_2.M_F_COD = TAB_FUNCIONARIO.F_Codigo  AND SQL_COMP_1.M_FERIAS <> SQL_COMP_2.M_FERIAS AND (SQL_COMP_1.Mes_Ano = '" & txt_1 & "') AND (SQL_COMP_2.Mes_Ano = '" & txt_2 & "')"
End Sub

Private Sub Op4_Click()
    w_Sql = "SELECT 'OBS' as OP, SQL_COMP_1.M_OBS AS Desc_1, SQL_COMP_2.M_OBS AS Desc_2, SQL_COMP_1.Mes_Ano AS Per1, SQL_COMP_2.Mes_Ano AS Per2,TAB_FUNCIONARIO.F_Cod_L AS L, TAB_FUNCIONARIO.F_NOME AS N FROM SQL_COMP_1, SQL_COMP_2, TAB_FUNCIONARIO WHERE SQL_COMP_1.M_F_COD = SQL_COMP_2.M_F_COD AND SQL_COMP_2.M_F_COD = TAB_FUNCIONARIO.F_Codigo  AND SQL_COMP_1.M_OBS <> SQL_COMP_2.M_OBS AND (SQL_COMP_1.Mes_Ano = '" & txt_1 & "') AND (SQL_COMP_2.Mes_Ano = '" & txt_2 & "')"
End Sub

Private Sub Op5_Click()
    w_Sql = "SELECT '@' as OP, SQL_COMP_1.M_DT_ADM AS Desc_1, SQL_COMP_2.M_DT_ADM AS Desc_2, SQL_COMP_1.Mes_Ano AS Per1, SQL_COMP_2.Mes_Ano AS Per2, TAB_FUNCIONARIO.F_Cod_L AS L, TAB_FUNCIONARIO.F_NOME AS N FROM SQL_COMP_1, SQL_COMP_2, TAB_FUNCIONARIO WHERE SQL_COMP_1.M_F_COD = SQL_COMP_2.M_F_COD AND SQL_COMP_2.M_F_COD = TAB_FUNCIONARIO.F_Codigo  AND Format(SQL_COMP_1.M_DT_ADM,'dd/mm/yyyy') <> Format(SQL_COMP_2.M_DT_ADM,'dd/mm/yyyy') AND (SQL_COMP_1.Mes_Ano = '" & txt_1 & "') AND (SQL_COMP_2.Mes_Ano = '" & txt_2 & "')"
End Sub

Private Sub Op6_Click()
    w_Sql = "SELECT '(R)' as OP, SQL_COMP_1.M_DT_REG AS Desc_1, SQL_COMP_2.M_DT_REG AS Desc_2, SQL_COMP_1.Mes_Ano AS Per1, SQL_COMP_2.Mes_Ano AS Per2, TAB_FUNCIONARIO.F_Cod_L AS L, TAB_FUNCIONARIO.F_NOME AS N FROM SQL_COMP_1, SQL_COMP_2, TAB_FUNCIONARIO WHERE SQL_COMP_1.M_F_COD = SQL_COMP_2.M_F_COD AND SQL_COMP_2.M_F_COD = TAB_FUNCIONARIO.F_Codigo  AND Format(SQL_COMP_1.M_DT_REG,'dd/mm/yyyy') <> Format(SQL_COMP_2.M_DT_REG,'dd/mm/yyyy') AND (SQL_COMP_1.Mes_Ano = '" & txt_1 & "') AND (SQL_COMP_2.Mes_Ano = '" & txt_2 & "')"
End Sub

Private Sub Op7_Click()
    w_Sql = "SELECT '(D)' as OP, SQL_COMP_1.M_DT_DEM AS Desc_1, SQL_COMP_2.M_DT_DEM AS Desc_2, SQL_COMP_1.Mes_Ano AS Per1, SQL_COMP_2.Mes_Ano AS Per2, TAB_FUNCIONARIO.F_Cod_L AS L, TAB_FUNCIONARIO.F_NOME AS N FROM SQL_COMP_1, SQL_COMP_2, TAB_FUNCIONARIO WHERE SQL_COMP_1.M_F_COD = SQL_COMP_2.M_F_COD AND SQL_COMP_2.M_F_COD = TAB_FUNCIONARIO.F_Codigo AND Format(SQL_COMP_1.M_DT_DEM,'dd/mm/yyyy') <> Format(SQL_COMP_2.M_DT_DEM,'dd/mm/yyyy') AND (SQL_COMP_1.Mes_Ano = '" & txt_1 & "') AND (SQL_COMP_2.Mes_Ano = '" & txt_2 & "')"
End Sub

Private Sub txt_1_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub txt_2_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{tab}"

End Sub
