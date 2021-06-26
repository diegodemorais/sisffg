VERSION 5.00
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frmRelatorioComissao 
   Caption         =   "Relatório de Comissão"
   ClientHeight    =   1140
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   1140
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Outros "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   5655
      Begin Skin_Button.ctr_Button cmdFixosSaldos 
         Height          =   525
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "FIXOS SALDOS"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRelatorioComissao.frx":0000
         PICN            =   "frmRelatorioComissao.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0FF&
      Caption         =   " Comissão "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin Skin_Button.ctr_Button cmdSalarioGerente 
         Height          =   525
         Left            =   3720
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "GERENTES"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRelatorioComissao.frx":27CE
         PICN            =   "frmRelatorioComissao.frx":27EA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button cmdSalarioCX 
         Height          =   525
         Left            =   1920
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "CAIXAS"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRelatorioComissao.frx":4F9C
         PICN            =   "frmRelatorioComissao.frx":4FB8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button cmdComissaoVendedor 
         Height          =   525
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "VENDEDORES"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRelatorioComissao.frx":776A
         PICN            =   "frmRelatorioComissao.frx":7786
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
End
Attribute VB_Name = "frmRelatorioComissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdComissaoVendedor_Click()
    Hide
    'If adoReg.Recordset.Fields("M_MES") = "1" Or adoReg.Recordset.Fields("M_MES") = "2" Or adoReg.Recordset.Fields("M_MES") = "3" Then
    '    If de.rscmdRptComissTMPFixo_Grouping.State = 1 Then de.rscmdRptComissTMPFixo_Grouping.Close
    '    de.cmdRptComissTMPFixo_Grouping
    '    rptComissMwtsFixo.Show
    'Else
        If de.rscmdRptComissTMP_Grouping.State = 1 Then de.rscmdRptComissTMP_Grouping.Close
        de.cmdRptComissTMP_Grouping
        rptComissMwts.Show
    'End If
End Sub

Private Sub cmdFixosSaldos_Click()
Dim mes, ano As String
    
    Hide
    If de.rscmdSqlFixosSaldos_Grouping.State = 1 Then de.rscmdSqlFixosSaldos_Grouping.Close
        
    'frm_ESCOLHA_DATA.Show 1
    
    'dtIni = CDate(frm_ESCOLHA_DATA.TXT_DT_INICIAL)
    'dtFim = CDate(frm_ESCOLHA_DATA.TXT_DT_FINAL)
    
     ano = InputBox("Entre com o Ano:")
     mes = InputBox("Entre com o Mês:")

    If ano <> "" And mes <> "" Then
        de.cmdSqlFixosSaldos_Grouping ano, mes
    
        rptFixosSaldos.Sections("secCab").Controls("lbData").Caption = "(" & mes & "/" & ano & ")"
    
        'rptSalarioGerentes.Sections("SecCab").Controls("lbTitulo").Caption = "SAL. G (" & Month(dtIni) & ")"
 
        rptFixosSaldos.Show
    End If
End Sub

Private Sub cmdSalarioCX_Click()
    Hide
    If de.rscmdSqlSalarioCxNOVO.State = 1 Then de.rscmdSqlSalarioCxNOVO.Close
     
    
    de.cmdSqlSalarioCxNOVO frm_Alt_Fic_Mensal_VIS.TXT_ANO, frm_Alt_Fic_Mensal_VIS.TXT_MES
    rptSalarioCxNOVO.Sections("SecCab").Controls("lbTitulo").Caption = "SAL. CXs. (" & frm_Alt_Fic_Mensal_VIS.TXT_MES & ")"
    
         
    rptSalarioCxNOVO.Show
End Sub

Private Sub cmdSalarioGerente_Click()
Dim dtIni, dtFim As Date
    
'If de.rscmdSqlSalarioGerentes.State = 1 Then de.rscmdSqlSalarioGerentes.Close
    
    frm_ESCOLHA_DATA.Show 1
    
    dtIni = CDate(frm_ESCOLHA_DATA.TXT_DT_INICIAL)
    dtFim = CDate(frm_ESCOLHA_DATA.TXT_DT_FINAL)
    
     'dtIni = InputBox("Entre com a Data Inicial:", , Format(Date, "DD/MM/YYYY"))
     'dtFim = InputBox("Entre com a Data Final:", , Format(Date, "DD/MM/YYYY"))
     
    Hide
    If de.rscmdSqlSalarioGerentes.State = 1 Then de.rscmdSqlSalarioGerentes.Close
     
    de.cmdSqlSalarioGerentes dtIni, dtFim
    rptSalarioGerentes.Sections("SecCab").Controls("lbTitulo").Caption = "SAL. G (" & Month(dtIni) & ")"
    'rptSalarioGerentes.Sections("SecCab").Controls("lbData").Caption = Format(Date, "DD=MM") & " " & Format(Time, "hh=mm")
         
    rptSalarioGerentes.Show
    
End Sub
