VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{9A4D18F7-4EC7-11D5-9E33-0040C78773FC}#1.0#0"; "VBxPOLITEC.ocx"
Begin VB.Form frm_Import 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importção da Central"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4470
   Icon            =   "frm_Import.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VBXPolitec.ocxProgressBarTexto pBar 
      Height          =   315
      Left            =   600
      TabIndex        =   11
      Top             =   2340
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorFundo  =   -2147483643
      MaxProgress     =   50
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "&Importar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1800
      TabIndex        =   9
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdImport 
      Height          =   615
      Index           =   0
      Left            =   1080
      Picture         =   "frm_Import.frx":1042
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   735
   End
   Begin VB.Frame painel 
      Caption         =   "Opções :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   1080
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
      Begin VB.CheckBox CK 
         Caption         =   "Vendas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox CK 
         Caption         =   "Emp."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox CK 
         Caption         =   "(B)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox CK 
         Caption         =   "Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox CK 
         Caption         =   "Crediários"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   1482
      ButtonWidth     =   1376
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Fechar"
            Key             =   "fechar"
            Object.ToolTipText     =   "Fechar (Alt+F)"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
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
            Picture         =   "frm_Import.frx":134C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Import.frx":1666
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Import.frx":1980
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Import.frx":1C9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Import.frx":1FB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Import.frx":22CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Importação da Central"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   90
      TabIndex        =   7
      Top             =   885
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Importação da Central"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   4215
   End
End
Attribute VB_Name = "frm_Import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BarraF_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err1
   
    Select Case Button.key
        Case "fechar": Fechar

    End Select

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub Fechar()
On Error GoTo err1
    Hide
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub

Private Sub cmdImport_Click(Index As Integer)
Dim w_Access As Access.Application
Set w_Access = New Access.Application

On Error GoTo ErrOpen

Inicio:

de.cnc.Close
de.cncDBase.Close
w_Access.OpenCurrentDatabase strDirBase, True
w_Access.Visible = False

GoTo 1
ErrOpen:
        
    MsgBox Err.Number & " : " & Err.Description, vbCritical


1:

    wQtde = 0
    If CK(0).value = 1 Then wQtde = wQtde + 1
    If CK(1).value = 1 Then wQtde = wQtde + 1
    If CK(2).value = 1 Then wQtde = wQtde + 1
    If CK(3).value = 1 Then wQtde = wQtde + 1
    If CK(4).value = 1 Then wQtde = wQtde + 1
    
    If wQtde > 0 Then w_Soma = Int(50 / wQtde)
    painel.Visible = False
    Pause 1
    pBar.Visible = True
    pBar.value = 1
    
On Error Resume Next
    '*** Crediários ***
    If CK(0).value = 1 Then
    '*** Deleta os Arquivos , depois copias os arquivos da central p/ o Diretorio do sistema
        Kill strDirBaseCentral & "\lojb08*.*"
            
        pBar.Text = "Crediários - Lojb081"
        CopiarArq strDirBaseServer & "\Lojb081.*", strDirBaseCentral
        pBar.value = w_Soma / 2
        
        pBar.Text = "Crediários - Lojb082"
        CopiarArq strDirBaseServer & "\Lojb082.*", strDirBaseCentral
        pBar.value = pBar.value + w_Soma / 2
    
    End If  '*** Crediários ***
    
On Error GoTo err1

    '*** Clientes ***
    If CK(1).value = 1 Then
    '*** Deleta a Tabela Se existir
    '*** Inporta a Tabela p/ o Access
    '*** Lojb011 ***
        pBar.Text = "Clientes - Lojb108"        '*** Deleta a Tabela se Existir
        w_Access.DoCmd.TransferDatabase acImport, "Paradox 7.x", strDirBaseServer, acTable, "lojb108", "Lojb108", False
        w_Access.DoCmd.RunSQL "DELETE From Lojb108"
        w_Access.DoCmd.RunSQL "INSERT INTO Lojb108 SELECT * FROM Lojb1081"
        DropTable w_Access, "Lojb1081"       '*** Deleta a Tabela se Existir
        
        pBar.value = pBar.value + w_Soma
    End If  '*** Clientes ***

    '*** Lojas ***
    If CK(2).value = 1 Then
    '*** Deleta a Tabela Se existir
    '*** Inporta a Tabela p/ o Access
    '*** Lojb011 ***
        pBar.Text = "Logos - Lojb010"
        DropTable w_Access, "Lojb010"       '*** Deleta a Tabela se Existir
        w_Access.DoCmd.TransferDatabase acImport, "dbase IV", strDirBaseServer, acTable, "lojb010", "Lojb010", False
        pBar.value = pBar.value + w_Soma
    End If  '*** Lojas ***
    
    '*** Funcionarios ***
    If CK(3).value = 1 Then
    '*** Deleta a Tabela Se existir
    '*** Inporta a Tabela p/ o Access
    '*** Lojb011 ***
        pBar.Text = "Emp. - Lojb011"
        w_Access.DoCmd.TransferDatabase acImport, "dbase IV", strDirBaseServer, acTable, "lojb011", "Lojb011", False
        w_Access.DoCmd.RunSQL "DELETE From Lojb011"
        w_Access.DoCmd.RunSQL "INSERT INTO Lojb011 SELECT * FROM Lojb0111"
        DropTable w_Access, "Lojb0111"       '*** Deleta a Tabela se Existir
        
        pBar.value = pBar.value + w_Soma
    End If  '*** Funcionarios ***
    
    
    '*** Vendas ***
    If CK(4).value = 1 Then
    '*** Deleta a Tabela Se existir
    '*** Inporta a Tabela p/ o Access
    '*** Lojb006 ***
        pBar.Text = "Vendas - Lojb006"
        w_Access.DoCmd.TransferDatabase acImport, "dbase IV", strDirBaseServer, acTable, "lojb006", "Lojb006", False
        w_Access.DoCmd.RunSQL "DELETE From Lojb006"
        w_Access.DoCmd.RunSQL "INSERT INTO Lojb006 SELECT LOJB0061.* FROM Lojb0061"
        DropTable w_Access, "Lojb0061"       '*** Deleta a Tabela se Existir
        pBar.value = pBar.value + (w_Soma / 3) / 5
            
    '*** Lojb015 ***
        pBar.Text = "Vendas - Lojb015"
        w_Access.DoCmd.TransferDatabase acImport, "dbase IV", strDirBaseServer, acTable, "lojb015", "Lojb015", False
        w_Access.DoCmd.RunSQL "DELETE From Lojb015"
        w_Access.DoCmd.RunSQL "DELETE From Lojb0151 WHERE (((DT_MOV)<=#" & Date - 45 & "#));"
        w_Access.DoCmd.RunSQL "INSERT INTO Lojb015 (LOJA, DT_MOV, CONTROLE, TIPO_DOC, NUMERO, OPERACAO, CAIXA, GERENTE, VENDEDOR, VND, DEV, DESCPCT, DESCONTO, DESCCORT, VALORTOT, PAGTO, FATOR, F_N, MOEDA1, VALOR1, MOEDA2, SIGLA, INDICE, LANCADO, VALORBRT, DEVOLUCAO, Cod_VND )" _
                            & " SELECT lojb0151.LOJA, lojb0151.DT_MOV, lojb0151.CONTROLE, lojb0151.TIPO_DOC, lojb0151.NUMERO, lojb0151.OPERACAO, lojb0151.CAIXA, lojb0151.GERENTE, lojb0151.VENDEDOR, lojb0151.VND, lojb0151.DEV, lojb0151.DESCPCT, lojb0151.DESCONTO, lojb0151.DESCCORT, lojb0151.VALORTOT, lojb0151.PAGTO, lojb0151.FATOR, lojb0151.F_N, lojb0151.MOEDA1, lojb0151.VALOR1, lojb0151.MOEDA2, lojb0151.SIGLA, lojb0151.INDICE, lojb0151.LANCADO, lojb0151.VALORBRT, lojb0151.DEVOLUCAO, [LOJA] & [dt_mov] & [controle] AS Codigo FROM lojb0151;"
        
        DropTable w_Access, "Lojb0151"       '*** Deleta a Tabela se Existir
        
        pBar.value = pBar.value + (w_Soma / 3) / 5
            
    '*** Lojb016 ***
        pBar.Text = "Vendas - Lojb016"
        w_Access.DoCmd.TransferDatabase acImport, "dbase IV", strDirBaseServer, acTable, "lojb016", "Lojb016", False
        w_Access.DoCmd.RunSQL "DELETE From Lojb016"
        w_Access.DoCmd.RunSQL "DELETE From Lojb0161 WHERE (((DT_MOV)<=#" & Date - 45 & "#));"
        w_Access.DoCmd.RunSQL "INSERT INTO Lojb016 ( LOJA, DT_MOV, CONTROLE, TERMINAL, CODIGO, TRANSACAO, VR, QT, TIPOIPI, DESCONTO, LANCADO, VALOR, ICMS, TIPOICMS, IPI, CT, TI, Cod_VND )" _
                            & " SELECT lojb0161.LOJA, lojb0161.DT_MOV, lojb0161.CONTROLE, lojb0161.TERMINAL, lojb0161.CODIGO, lojb0161.TRANSACAO, lojb0161.VR, lojb0161.QT, lojb0161.TIPOIPI, lojb0161.DESCONTO, lojb0161.LANCADO, lojb0161.VALOR, lojb0161.ICMS, lojb0161.TIPOICMS, lojb0161.IPI, lojb0161.CT, lojb0161.TI, [LOJA] & [dt_mov] & [controle] AS Cod_Vnd FROM lojb0161;"

        DropTable w_Access, "Lojb0161"       '*** Deleta a Tabela se Existir
        
        pBar.value = pBar.value + (w_Soma / 3) / 5
    
    '*** Lojb022 ***
        pBar.Text = "Vendas - Lojb022"
        w_Access.DoCmd.TransferDatabase acImport, "dbase IV", strDirBaseServer, acTable, "lojb022", "Lojb022", False
        w_Access.DoCmd.RunSQL "DELETE From Lojb022"
        w_Access.DoCmd.RunSQL "DELETE From Lojb0221 WHERE (((DT_lanc)<=#" & Date - 45 & "#));"
        w_Access.DoCmd.RunSQL "INSERT INTO LOJB022 ( LOJA, DT_LANC, CONTROLE, DT_MOV, TIPO, VALOR, CRR, LANCADO, PARCELA, COND_PGTO, COD_VND )" _
                            & "SELECT lojb0221.LOJA, lojb0221.DT_LANC, lojb0221.CONTROLE, lojb0221.DT_MOV, lojb0221.TIPO, lojb0221.VALOR, lojb0221.CRR, lojb0221.LANCADO, lojb0221.PARCELA, lojb0221.COND_PGTO, [LOJA] & [DT_LANC] & [controle] AS COD_VND FROM lojb0221;"
        DropTable w_Access, "Lojb0221"       '*** Deleta a Tabela se Existir
        
        pBar.value = pBar.value + (w_Soma / 3) / 5
    
    '*** Lojb135 ***
        pBar.Text = "Vendas - Lojb135"
        w_Access.DoCmd.TransferDatabase acImport, "dbase IV", strDirBaseServer, acTable, "lojb135", "Lojb135", False
        w_Access.DoCmd.RunSQL "DELETE From Lojb135"
        w_Access.DoCmd.RunSQL "INSERT INTO Lojb135 SELECT * FROM Lojb1351"
        DropTable w_Access, "Lojb1351"       '*** Deleta a Tabela se Existir
        
        pBar.value = pBar.value + (w_Soma / 3) / 5
    End If  '*** Vendas ***


    CK(0).value = 0
    CK(1).value = 0
    CK(2).value = 0
    CK(3).value = 0
    CK(4).value = 0


    MsgBox "Importação concluída com sucesso!" & Chr(13) & "Será feito um reparo no Banco de Dados!" & Chr(13) & "Certifique-se, de que não existe nenhuma máquina com o Sistema aberto!", vbExclamation
    
    w_Access.CloseCurrentDatabase '*** Fecha a Conexão com o Banco
    '*** Compacta BD
    pBar.Width = pBar.Width + 700
    pBar.Left = pBar.Left - 350
    pBar.Text = "Reparando e Compactando o Banco de dados !"
    
    Pause 0.1
    pBar.value = 10
    w_Access.DBEngine.CompactDatabase strDirBase, strDirBaseCentral & "\Banco_C.mdb"
    pBar.value = 30
    Kill strDirBase  '*** Exclu o Atual
    pBar.value = 40
    CopiarArq strDirBaseCentral & "\Banco_C.mdb", strDirBase '*** copia o Compactado com o Nome certo
    Kill strDirBaseCentral & "\Banco_C.mdb" 'Exclui Banco Compactado
    pBar.value = 50
    
    Set w_Access = Nothing  'libera a memoria
    
    Unload Me
    MsgBox "Você deverá entrar no sistema novamente!", vbInformation
    End

sair:
    
    Exit Sub
err1:
    If Err.Number = 3044 Then
        MsgBox "O SERVIDOR DA CENTRAL DE ONDE SERÃO IMPORTADAS AS TABELAS, DEVE ESTAR DESLIGADO OU O CAMINHO ESTA INCORRETO!", vbCritical
    Else
        MsgBox Err.Number & " : " & Err.Description, vbCritical
    End If
    Resume sair
End Sub



Sub DropTable(ByRef w_Access, strTableName As String)
On Error GoTo err1

     w_Access.DoCmd.RunSQL "Drop Table " & strTableName     '*** Deleta a Tabela se Existir

sair:
    Exit Sub
err1:
    MsgBox Err.Description, vbCritical
    Resume sair
End Sub


