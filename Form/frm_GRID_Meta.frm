VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form frm_GRID_Meta 
   Caption         =   "Cadastro de Metas"
   ClientHeight    =   11190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   11190
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLiq 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$"" #.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5970
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   10005
      Width           =   1500
   End
   Begin VB.TextBox txtMeta 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$"" #.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "R$ 0,00"
      Top             =   10365
      Width           =   1500
   End
   Begin MSDataGridLib.DataGrid Grid 
      Bindings        =   "frm_GRID_Meta.frx":0000
      Height          =   8640
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   15240
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   18
      TabAction       =   1
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "MT_F_LOJA"
         Caption         =   "LOJA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "MT_MES"
         Caption         =   "MÊS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "MT_ANO"
         Caption         =   "ANO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "MT_VALOR"
         Caption         =   "META"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt_Ano 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.ComboBox txt_Mes 
      Height          =   315
      ItemData        =   "frm_GRID_Meta.frx":0015
      Left            =   840
      List            =   "frm_GRID_Meta.frx":003D
      TabIndex        =   0
      Top             =   480
      Width           =   690
   End
   Begin VB.CommandButton cmdPesq 
      Caption         =   "&Buscar"
      Height          =   735
      Left            =   3960
      Picture         =   "frm_GRID_Meta.frx":0068
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   735
   End
   Begin MSAdodcLib.Adodc adoReg 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   10860
      Width           =   5400
      _ExtentX        =   9525
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
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Meta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3540
      TabIndex        =   8
      Top             =   10005
      Width           =   1155
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   345
      Left            =   3285
      Top             =   9960
      Width           =   1620
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Piso Líq"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   6285
      TabIndex        =   9
      Top             =   9645
      Width           =   840
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   465
      Left            =   3285
      Top             =   10320
      Width           =   1620
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano:"
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
      Left            =   2040
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mês:"
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
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   0
      Top             =   120
      Width           =   5370
   End
End
Attribute VB_Name = "frm_GRID_Meta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPesq_Click()

    If de.rscmdSqlGridMeta2.State = 1 Then de.rscmdSqlGridMeta2.Close
    de.cmdSqlGridMeta2
    
    Set adoReg.Recordset = de.rscmdSqlGridMeta2.Clone

    If TXT_ANO = "" Then TXT_ANO = Year(Now())
    If TXT_MES < 1 Or TXT_MES > 12 Then TXT_MES = Month(Now())
    
    adoReg.Recordset.Filter = "MT_MES = '" & TXT_MES & "' AND MT_ANO = '" & TXT_ANO & "'"
    
    Call Calc_Total_Meta
    

End Sub

Sub Calc_Total_Meta()

   
    Dim adoGrid As ADODB.Recordset
    Dim wMeta

    wMeta = 0
    
    Set adoGrid = adoReg.Recordset.Clone
    adoGrid.Filter = "MT_MES = '" & TXT_MES & "' AND MT_ANO = '" & TXT_ANO & "'"
    
    If adoGrid.RecordCount <> 0 Then adoGrid.MoveFirst
    Do While Not adoGrid.EOF
        wMeta = wMeta + CDbl(adoGrid.Fields("MT_VALOR"))

        adoGrid.MoveNext
    Loop
    
    txtMeta = Format(wMeta, "R$ #,##0.00")

End Sub


Private Sub Form_Load()
      
TXT_MES = Month(Now())
TXT_ANO = Year(Now())

    If de.rscmdSqlGridMeta2.State = 1 Then de.rscmdSqlGridMeta2.Close
    de.cmdSqlGridMeta2
    
    Set adoReg.Recordset = de.rscmdSqlGridMeta2.Clone
    
    adoReg.Recordset.Filter = "MT_MES = '" & TXT_MES & "' AND MT_ANO = '" & TXT_ANO & "'"
    
    Call Calc_Total_Meta

End Sub

Private Sub Grid_AfterUpdate()
    Call Calc_Total_Meta
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Sendkeys "{tab}"
    End If
End Sub


Private Sub TXT_ANO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdPesq_Click
    End If
End Sub

Private Sub TXT_MES_Change()
    cmdPesq_Click
End Sub
