VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_Cad_Emprestimo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Empréstimo"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "frm_Cad_Emprestimo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt_Valor_Parc 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1755
      TabIndex        =   16
      Top             =   2640
      Width           =   1200
   End
   Begin MSDataListLib.DataCombo dcFunc 
      Bindings        =   "frm_Cad_Emprestimo.frx":1CFA
      Height          =   315
      Left            =   2520
      TabIndex        =   15
      Top             =   1320
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "F_NOME"
      BoundColumn     =   "F_Codigo"
      Text            =   ""
      Object.DataMember      =   "TAB_FUNCIONARIO"
   End
   Begin VB.TextBox txtE_C_CODIGO 
      Height          =   285
      Left            =   6060
      TabIndex        =   14
      Top             =   3420
      Width           =   660
   End
   Begin VB.TextBox txtE_M_NFICHA 
      Height          =   285
      Left            =   6060
      TabIndex        =   12
      Top             =   3030
      Width           =   660
   End
   Begin VB.TextBox txtE_VALOR 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1860
      TabIndex        =   10
      Top             =   1800
      Width           =   1200
   End
   Begin VB.TextBox txtE_DT_PREV 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5325
      TabIndex        =   8
      Top             =   2250
      Width           =   1260
   End
   Begin VB.TextBox txtE_QTDE_PARC 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1875
      TabIndex        =   6
      Top             =   2265
      Width           =   660
   End
   Begin VB.TextBox txtE_JUROS 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   5
      EndProperty
      Height          =   285
      Left            =   5940
      TabIndex        =   4
      Top             =   1785
      Width           =   660
   End
   Begin VB.TextBox txtE_F_COD 
      Height          =   300
      Left            =   1860
      TabIndex        =   2
      Top             =   1320
      Width           =   660
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   1429
      ButtonWidth     =   1376
      ButtonHeight    =   1376
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
               Picture         =   "frm_Cad_Emprestimo.frx":1D0B
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Emprestimo.frx":2025
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Cad_Emprestimo.frx":3307
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Valor Parc."
      Height          =   195
      Index           =   2
      Left            =   945
      TabIndex        =   17
      Top             =   2685
      Width           =   780
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "E_C_CODIGO:"
      Height          =   255
      Index           =   8
      Left            =   4695
      TabIndex        =   13
      Top             =   3465
      Width           =   1335
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "E_M_NFICHA:"
      Height          =   255
      Index           =   7
      Left            =   4695
      TabIndex        =   11
      Top             =   3075
      Width           =   1335
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      Height          =   195
      Index           =   5
      Left            =   1425
      TabIndex        =   9
      Top             =   1845
      Width           =   405
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Dt. 1º Parcela :"
      Height          =   195
      Index           =   4
      Left            =   4230
      TabIndex        =   7
      Top             =   2295
      Width           =   1080
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Qtde de Parcelas :"
      Height          =   195
      Index           =   3
      Left            =   525
      TabIndex        =   5
      Top             =   2310
      Width           =   1320
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "%:"
      Height          =   195
      Index           =   1
      Left            =   5715
      TabIndex        =   3
      Top             =   1830
      Width           =   165
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Emp.:"
      Height          =   195
      Index           =   0
      Left            =   1425
      TabIndex        =   1
      Top             =   1380
      Width           =   405
   End
End
Attribute VB_Name = "frm_Cad_Emprestimo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
