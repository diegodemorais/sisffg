VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_Alt_Fic_Mensal_Visualizar_Dupla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VISUALIZAÇÃO - DUPLA"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   12000
   Icon            =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   12000
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TXT_CAMPOS 
      BackColor       =   &H0080FFFF&
      DataField       =   "F_NOME"
      DataSource      =   "ADOREG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   8055
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2160
      Width           =   3720
   End
   Begin VB.TextBox TXT_CAMPOS 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      DataField       =   "Logo"
      DataSource      =   "ADOREG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   6165
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1785
      Width           =   855
   End
   Begin VB.TextBox txtM_MES 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      DataField       =   "data"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MMMM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      DataSource      =   "ADOREG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1785
      Width           =   1620
   End
   Begin VB.TextBox TXT_ANO 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      DataField       =   "M_ANO"
      DataSource      =   "ADOREG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   8820
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1785
      Width           =   660
   End
   Begin VB.TextBox TXT_CAMPOS 
      DataField       =   "M_FERIAS"
      DataSource      =   "ADOREG"
      Height          =   525
      Index           =   4
      Left            =   7200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Top             =   2520
      Width           =   4575
   End
   Begin VB.TextBox TXT_CAMPOS 
      DataField       =   "M_OBS"
      DataSource      =   "ADOREG"
      Height          =   525
      Index           =   5
      Left            =   7200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Top             =   3120
      Width           =   4575
   End
   Begin VB.TextBox TXT_CAMPOS 
      DataField       =   "M_ANOTACAO"
      DataSource      =   "ADOREG"
      Height          =   525
      Index           =   6
      Left            =   7200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Top             =   3720
      Width           =   4575
   End
   Begin VB.TextBox TXT_CAMPOS 
      Alignment       =   2  'Center
      DataField       =   "M_DT_ADM"
      DataSource      =   "ADOREG"
      Height          =   285
      Index           =   2
      Left            =   9675
      TabIndex        =   26
      Top             =   1770
      Width           =   990
   End
   Begin VB.TextBox TXT_CAMPOS 
      Alignment       =   2  'Center
      DataField       =   "M_DT_REG"
      DataSource      =   "ADOREG"
      Height          =   285
      Index           =   3
      Left            =   10800
      TabIndex        =   25
      Top             =   1770
      Width           =   960
   End
   Begin VB.TextBox TXT_MAIS 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "R$ #.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   6510
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7410
      Width           =   1140
   End
   Begin VB.TextBox TXT_MENOS 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "R$ #.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   300
      Left            =   8190
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7410
      Width           =   1140
   End
   Begin VB.TextBox TXT_TOTAL 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "R$ #.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10620
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7410
      Width           =   1140
   End
   Begin VB.TextBox txt_Cred 
      Alignment       =   2  'Center
      DataField       =   "Cred"
      DataSource      =   "ADOREG"
      Height          =   285
      Left            =   6150
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox TXT_CAMPOS 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      DataField       =   "M_F_COD"
      DataSource      =   "ADOREG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   7
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2160
      Width           =   855
   End
   Begin VB.ComboBox txt_Mes1 
      Height          =   315
      ItemData        =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":12D2
      Left            =   7020
      List            =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":12FA
      TabIndex        =   17
      Top             =   1095
      Width           =   660
   End
   Begin VB.ComboBox txt_Ano1 
      Height          =   315
      ItemData        =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":1325
      Left            =   7950
      List            =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":134A
      TabIndex        =   16
      Text            =   "2005"
      Top             =   1095
      Width           =   750
   End
   Begin VB.ComboBox txt_Ano2 
      Height          =   315
      ItemData        =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":1390
      Left            =   2085
      List            =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":13B5
      TabIndex        =   15
      Text            =   "2005"
      Top             =   1110
      Width           =   750
   End
   Begin VB.ComboBox txt_Mes2 
      Height          =   315
      ItemData        =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":13FB
      Left            =   1155
      List            =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":1423
      TabIndex        =   14
      Top             =   1110
      Width           =   660
   End
   Begin VB.TextBox TXT_CAMPOS 
      BackColor       =   &H00800000&
      DataField       =   "F_NOME"
      DataSource      =   "adoReg2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   15
      Left            =   2145
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2145
      Width           =   3720
   End
   Begin VB.TextBox TXT_CAMPOS 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      DataField       =   "Logo"
      DataSource      =   "adoReg2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   14
      Left            =   270
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1770
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      DataField       =   "data"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MMMM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      DataSource      =   "adoReg2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1770
      Width           =   1620
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      DataField       =   "M_ANO"
      DataSource      =   "adoReg2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2910
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1770
      Width           =   660
   End
   Begin VB.TextBox TXT_CAMPOS 
      DataField       =   "M_FERIAS"
      DataSource      =   "adoReg2"
      Enabled         =   0   'False
      Height          =   525
      Index           =   13
      Left            =   1290
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2505
      Width           =   4575
   End
   Begin VB.TextBox TXT_CAMPOS 
      DataField       =   "M_OBS"
      DataSource      =   "adoReg2"
      Enabled         =   0   'False
      Height          =   525
      Index           =   12
      Left            =   1290
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3105
      Width           =   4575
   End
   Begin VB.TextBox TXT_CAMPOS 
      DataField       =   "M_ANOTACAO"
      DataSource      =   "adoReg2"
      Enabled         =   0   'False
      Height          =   525
      Index           =   11
      Left            =   1290
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   3705
      Width           =   4575
   End
   Begin VB.TextBox TXT_CAMPOS 
      Alignment       =   2  'Center
      DataField       =   "M_DT_ADM"
      DataSource      =   "adoReg2"
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   3810
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1755
      Width           =   975
   End
   Begin VB.TextBox TXT_CAMPOS 
      Alignment       =   2  'Center
      DataField       =   "M_DT_REG"
      DataSource      =   "adoReg2"
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   4890
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1755
      Width           =   960
   End
   Begin VB.TextBox TXT_MAIS2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "R$ #.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7425
      Width           =   1140
   End
   Begin VB.TextBox TXT_MENOS2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "R$ #.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   300
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7425
      Width           =   1140
   End
   Begin VB.TextBox TXT_TOTAL2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "R$ #.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7425
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "Cred"
      DataSource      =   "adoReg2"
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2745
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox TXT_CAMPOS 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      DataField       =   "M_F_COD"
      DataSource      =   "adoReg2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   8
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2145
      Width           =   855
   End
   Begin Skin_Button.ctr_Button bt_OK1 
      Height          =   345
      Left            =   8715
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1080
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":144E
      PICN            =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":146A
      PICH            =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":18BC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc ADOREG 
      Height          =   330
      Left            =   6225
      Top             =   6705
      Visible         =   0   'False
      Width           =   5250
      _ExtentX        =   9260
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
      Caption         =   "REGISTRO : 0/0"
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
   Begin MSAdodcLib.Adodc adoConta 
      Height          =   330
      Left            =   7350
      Top             =   7065
      Visible         =   0   'False
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   2
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
      Caption         =   "REGISTRO : 0/0"
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
   Begin MSDataGridLib.DataGrid grid_Conta 
      Bindings        =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":1BD6
      Height          =   3045
      Left            =   6150
      TabIndex        =   20
      Top             =   4335
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   5371
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "CONTAS"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "C_DT"
         Caption         =   "DATA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "TP_DESC"
         Caption         =   "CONTAS"
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
         DataField       =   "TP_DESC"
         Caption         =   "CONTA"
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
         DataField       =   "C_VALOR"
         Caption         =   "VALOR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "R$ #0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "c_tp_op"
         Caption         =   "OP."
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
      BeginProperty Column05 
         DataField       =   "c_Visto"
         Caption         =   "Visto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "ok"
            FalseValue      =   "não"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "C_CODIGO"
         Caption         =   "COD"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   2129,953
         EndProperty
         BeginProperty Column02 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   2280,189
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   374,74
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column06 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   0
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":1BED
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":1F07
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":2221
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":253B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":2855
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":2B6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":2E89
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoConta2 
      Height          =   330
      Left            =   1440
      Top             =   7185
      Visible         =   0   'False
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   2
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
      Caption         =   "REGISTRO : 0/0"
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
   Begin MSAdodcLib.Adodc adoReg2 
      Height          =   330
      Left            =   240
      Top             =   6885
      Visible         =   0   'False
      Width           =   5730
      _ExtentX        =   10107
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
      Caption         =   "REGISTRO : 0/0"
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
   Begin MSDataGridLib.DataGrid grid_conta2 
      Bindings        =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":32DB
      Height          =   3045
      Left            =   225
      TabIndex        =   34
      Top             =   4350
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   5371
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      Enabled         =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "CONTAS"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "C_DT"
         Caption         =   "DATA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "TP_DESC"
         Caption         =   "CONTAS"
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
         DataField       =   "C_DESC"
         Caption         =   "DESCRIÇÃO"
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
         DataField       =   "C_VALOR"
         Caption         =   "VALOR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "R$ #0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "c_tp_op"
         Caption         =   "OP."
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
      BeginProperty Column05 
         DataField       =   "c_Visto"
         Caption         =   "Visto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "ok"
            FalseValue      =   "não"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "C_CODIGO"
         Caption         =   "COD"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   2129,953
         EndProperty
         BeginProperty Column02 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   2280,189
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   374,74
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column06 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   0
         EndProperty
      EndProperty
   End
   Begin Skin_Button.ctr_Button bt_OK2 
      Height          =   345
      Left            =   2880
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1095
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":32F3
      PICN            =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":330F
      PICH            =   "frm_Alt_Fic_Mensal_Visualizar_Dupla2.frx":3761
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.Toolbar BarraF 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   1429
      ButtonWidth     =   1852
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Fechar (F5)"
            Key             =   "fechar"
            Object.ToolTipText     =   "Fechar (Alt+F) ou (F5)"
            ImageIndex      =   1
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Editar"
            Key             =   "editar"
            Object.ToolTipText     =   "Editar Alteração (Alt+E)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   6150
      TabIndex        =   66
      Top             =   2190
      Width           =   555
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(B)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   6180
      TabIndex        =   65
      Top             =   1575
      Width           =   255
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MÊS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   7065
      TabIndex        =   64
      Top             =   1575
      Width           =   1635
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ANO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   8835
      TabIndex        =   63
      Top             =   1575
      Width           =   450
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(F)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   6150
      TabIndex        =   62
      Top             =   2520
      Width           =   240
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   6150
      TabIndex        =   61
      Top             =   3120
      Width           =   390
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ANOTAÇÃO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   6150
      TabIndex        =   60
      Top             =   3720
      Width           =   1020
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   10050
      TabIndex        =   59
      Top             =   1485
      Width           =   285
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "®"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   11115
      TabIndex        =   58
      Top             =   1500
      Width           =   210
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   10
      Left            =   6180
      TabIndex        =   57
      Top             =   7365
      Width           =   240
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   11
      Left            =   7830
      TabIndex        =   56
      Top             =   7320
      Width           =   150
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "=      TOTAL "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   9435
      TabIndex        =   55
      Top             =   7485
      Width           =   1260
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   7695
      X2              =   7875
      Y1              =   1350
      Y2              =   1125
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mês/Ano:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   6135
      TabIndex        =   54
      Top             =   1170
      Width           =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7155
      TabIndex        =   53
      Top             =   825
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "F4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8025
      TabIndex        =   52
      Top             =   825
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2175
      TabIndex        =   51
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1305
      TabIndex        =   50
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mês/Ano:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   270
      TabIndex        =   49
      Top             =   1185
      Width           =   840
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   1845
      X2              =   2025
      Y1              =   1395
      Y2              =   1170
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   240
      TabIndex        =   48
      Top             =   2175
      Width           =   555
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(B)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   255
      TabIndex        =   47
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MÊS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   1275
      TabIndex        =   46
      Top             =   1560
      Width           =   1635
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ANO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   3045
      TabIndex        =   45
      Top             =   1560
      Width           =   450
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(F)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   240
      TabIndex        =   44
      Top             =   2505
      Width           =   240
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   240
      TabIndex        =   43
      Top             =   3105
      Width           =   390
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ANOTAÇÃO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   240
      TabIndex        =   42
      Top             =   3705
      Width           =   1020
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   16
      Left            =   4170
      TabIndex        =   41
      Top             =   1485
      Width           =   285
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "®"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   15
      Left            =   5325
      TabIndex        =   40
      Top             =   1485
      Width           =   210
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   14
      Left            =   270
      TabIndex        =   39
      Top             =   7365
      Width           =   240
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   13
      Left            =   1890
      TabIndex        =   38
      Top             =   7305
      Width           =   150
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "=      TOTAL "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   3510
      TabIndex        =   37
      Top             =   7470
      Width           =   1260
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   7860
      Left            =   5970
      Top             =   60
      Width           =   60
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   120
      Top             =   1080
      Width           =   3180
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   6330
      Left            =   120
      Top             =   1470
      Width           =   5835
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   6330
      Left            =   6045
      Top             =   1455
      Width           =   5835
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   6045
      Top             =   1065
      Width           =   3120
   End
   Begin VB.Menu mnu 
      Caption         =   "Menu"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Editar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuIR 
         Caption         =   "Ir para"
         Begin VB.Menu mnuf1 
            Caption         =   "(F1)"
            Shortcut        =   {F1}
         End
         Begin VB.Menu mnuF2 
            Caption         =   "(F2)"
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnuF3 
            Caption         =   "(F3)"
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuF4 
            Caption         =   "(F4)"
            Shortcut        =   {F4}
         End
      End
      Begin VB.Menu mnusep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFechar 
         Caption         =   "&Fechar"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "frm_Alt_Fic_Mensal_Visualizar_Dupla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W_LD_FILTRO As Boolean
Dim W_POS As Long
Dim w_Move As Boolean
Dim W_Pv  As Boolean ' SE É A PRIMEIRA VEZ QUE ENTREA NA TELA
Dim W_FILTRO As String
Dim W_INDEX As Byte
Dim w_Hab As Boolean


Private Sub Total()
Dim ADO_TOTAL As ADODB.Recordset

On Error GoTo err1
    
    TXT_MAIS = 0
    TXT_MENOS = 0
    TXT_TOTAL = 0
    
    Set ADO_TOTAL = adoConta.Recordset.Clone
    
    If Not ADO_TOTAL.EOF Then
        ADO_TOTAL.MoveFirst
        Do While Not ADO_TOTAL.EOF
            If ADO_TOTAL.Fields("C_valor") >= 0 And ADO_TOTAL.Fields("C_Tp_OP") = "+" Then
                TXT_MAIS = CDbl(TXT_MAIS) + ADO_TOTAL.Fields("C_VALOR")
            ElseIf ADO_TOTAL.Fields("C_valor") < 0 And ADO_TOTAL.Fields("C_Tp_OP") = "-" Then
                TXT_MENOS = CDbl(TXT_MENOS) + ADO_TOTAL.Fields("C_VALOR")
            End If
            ADO_TOTAL.MoveNext
        Loop
        
        TXT_TOTAL = CDbl(TXT_MAIS) - CDbl(TXT_MENOS)
    End If
    
    TXT_TOTAL = Format(CDbl(TXT_MENOS) + CDbl(TXT_MAIS), "R$ 0.00")
    TXT_MAIS = Format(TXT_MAIS, "R$ #0.00")
    TXT_MENOS = Format(TXT_MENOS, "R$ #0.00")
    
    
    'muda cor do total
    If TXT_TOTAL < 0 Then
        TXT_TOTAL.ForeColor = vbRed
    Else
        TXT_TOTAL.ForeColor = vbBlue
    End If
    


sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub



Private Sub Total2()
Dim ADO_TOTAL As ADODB.Recordset

On Error GoTo err1
    
    TXT_MAIS2 = 0
    TXT_MENOS2 = 0
    TXT_TOTAL2 = 0
    
    Set ADO_TOTAL = adoConta2.Recordset.Clone
    
    If Not ADO_TOTAL.EOF Then
        ADO_TOTAL.MoveFirst
        Do While Not ADO_TOTAL.EOF
            If ADO_TOTAL.Fields("C_valor") >= 0 And ADO_TOTAL.Fields("C_Tp_OP") = "+" Then
                TXT_MAIS2 = CDbl(TXT_MAIS2) + ADO_TOTAL.Fields("C_VALOR")
            ElseIf ADO_TOTAL.Fields("C_valor") < 0 And ADO_TOTAL.Fields("C_Tp_OP") = "-" Then
                TXT_MENOS2 = CDbl(TXT_MENOS2) + ADO_TOTAL.Fields("C_VALOR")
            End If
            ADO_TOTAL.MoveNext
        Loop
        
        TXT_TOTAL2 = CDbl(TXT_MAIS2) - CDbl(TXT_MENOS2)
    End If
    
    TXT_TOTAL2 = Format(CDbl(TXT_MENOS2) + CDbl(TXT_MAIS2), "R$ 0.00")
    TXT_MAIS2 = Format(TXT_MAIS2, "R$ #0.00")
    TXT_MENOS2 = Format(TXT_MENOS2, "R$ #0.00")
    
    
    'muda cor do total
    If TXT_TOTAL2 < 0 Then
        TXT_TOTAL2.ForeColor = vbRed
    Else
        TXT_TOTAL2.ForeColor = vbBlue
    End If
    

sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub





Private Sub adoReg2_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1
    
 If w_Hab = True And Not adoReg2.Recordset.EOF Then
    
    grid_conta2.Visible = True
    
    Set adoConta2.Recordset = de.cnc.Execute("SELECT TAB_DESC_CALC.C_CODIGO, TAB_DESC_CALC.C_N_FICHA, TAB_DESC_CALC.C_DT, TAB_TP_CONTA.TP_DESC, TAB_DESC_CALC.C_TP_OP, TAB_DESC_CALC.C_VALOR, TAB_DESC_CALC.C_VISTO, TAB_DESC_CALC.C_DESC FROM TAB_DESC_CALC, TAB_TP_CONTA WHERE TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_DESC_CALC.C_N_FICHA = " & adoReg2.Recordset.Fields("M_Nficha") & ") ORDER BY TAB_TP_CONTA.TP_DESC, TAB_DESC_CALC.C_TP_OP DESC").Clone
    adoConta2.Refresh
   ' Pause 0.3
    Total2
    TXT_TOTAL2.Refresh
    grid_conta2.Visible = True

ElseIf adoReg2.Recordset.EOF Then
    grid_conta2.Visible = False
    TXT_MENOS2 = ""
    TXT_MAIS2 = ""
    TXT_TOTAL2 = ""
End If

sair:
    Exit Sub
err1:
    If Not Err.Number = -2147217885 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair

End Sub



Private Sub bt_Edit1_Click()
    For i = 2 To 6
        TXT_CAMPOS(i).Enabled = Not TXT_CAMPOS(i).Enabled
    Next i
    grid_Conta.Enabled = Not grid_Conta.Enabled
    bt_Edit1.UseGreyscale = Not bt_Edit1.UseGreyscale
End Sub

Private Sub bt_Edit2_Click()
    For i = 9 To 13
        TXT_CAMPOS(i).Enabled = Not TXT_CAMPOS(i).Enabled
    Next i
    grid_conta2.Enabled = Not grid_conta2.Enabled
    bt_Edit2.UseGreyscale = Not bt_Edit2.UseGreyscale
End Sub

Private Sub bt_OK1_Click()
    w_Hab = False
    If de.rscmdSqlVisualizarFichas.State = 1 Then de.rscmdSqlVisualizarFichas.Close
    de.cmdSqlVisualizarFichas txt_Ano1, txt_Mes1
    Set ADOREG.Recordset = de.rscmdSqlVisualizarFichas.Clone          'de.cnc.Execute("SELECT TAB_FICHA_MENS.M_NFICHA, TAB_FUNCIONARIO.F_Cod_L AS LOGO, TAB_FICHA_MENS.M_ANO, TAB_FUNCIONARIO.F_NOME, TAB_FUNCIONARIO.F_DT_ADM, TAB_FUNCIONARIO.F_DT_REG, TAB_FICHA_MENS.M_FERIAS, TAB_FICHA_MENS.M_OBS, TAB_FUNCIONARIO.F_ANOTACAO, TAB_FICHA_MENS.M_TOTAL_MAIS AS MAIS, TAB_FICHA_MENS.M_TOTAL_MENOS AS MENOS, TAB_FICHA_MENS.M_TOTAL_MAIS - TAB_FICHA_MENS.M_TOTAL_MENOS AS TOTAL, '01/' + str(TAB_FICHA_MENS.M_MES) + '/' + str(TAB_FICHA_MENS.M_ANO) AS data, TAB_FICHA_MENS.M_MES as M_MES, TAB_FICHA_MENS.m_bloq as Bloq, Tab_Funcionario.F_Cod_Cred as Cred FROM TAB_FICHA_MENS, TAB_FUNCIONARIO WHERE TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo  Order By  TAB_FICHA_MENS.M_MES, TAB_FUNCIONARIO.F_Nome ").Clone
    de.rscmdSqlVisualizarFichas.Close

    If TXT_CAMPOS(7) <> "" And W_FILTRO <> "M_F_COD = " Then
        w_Hab = True
        ADOREG.Recordset.Filter = W_FILTRO
    Else
        grid_Conta.Visible = False
    End If
End Sub

Private Sub bt_OK1_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
    If KeyCode = 13 And Shift = 0 Then
        TXT_CAMPOS(2).SetFocus
    End If

End Sub

Private Sub bt_OK2_Click()
    w_Hab = False
    If de.rscmdSqlVisualizarFichas.State = 1 Then de.rscmdSqlVisualizarFichas.Close
    de.cmdSqlVisualizarFichas txt_Ano2, txt_Mes2
    Set adoReg2.Recordset = de.rscmdSqlVisualizarFichas.Clone          'de.cnc.Execute("SELECT TAB_FICHA_MENS.M_NFICHA, TAB_FUNCIONARIO.F_Cod_L AS LOGO, TAB_FICHA_MENS.M_ANO, TAB_FUNCIONARIO.F_NOME, TAB_FUNCIONARIO.F_DT_ADM, TAB_FUNCIONARIO.F_DT_REG, TAB_FICHA_MENS.M_FERIAS, TAB_FICHA_MENS.M_OBS, TAB_FUNCIONARIO.F_ANOTACAO, TAB_FICHA_MENS.M_TOTAL_MAIS AS MAIS, TAB_FICHA_MENS.M_TOTAL_MENOS AS MENOS, TAB_FICHA_MENS.M_TOTAL_MAIS - TAB_FICHA_MENS.M_TOTAL_MENOS AS TOTAL, '01/' + str(TAB_FICHA_MENS.M_MES) + '/' + str(TAB_FICHA_MENS.M_ANO) AS data, TAB_FICHA_MENS.M_MES as M_MES, TAB_FICHA_MENS.m_bloq as Bloq, Tab_Funcionario.F_Cod_Cred as Cred FROM TAB_FICHA_MENS, TAB_FUNCIONARIO WHERE TAB_FICHA_MENS.M_F_COD = TAB_FUNCIONARIO.F_Codigo  Order By  TAB_FICHA_MENS.M_MES, TAB_FUNCIONARIO.F_Nome ").Clone
    de.rscmdSqlVisualizarFichas.Close
    
    If TXT_CAMPOS(7) <> "" And W_FILTRO <> "M_F_COD = " Then
        w_Hab = True
        adoReg2.Recordset.Filter = W_FILTRO
    Else
        grid_Conta.Visible = False
    End If
End Sub

Private Sub bt_OK2_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift

End Sub

Private Sub Form_Activate()

    On Error Resume Next
    
On Error GoTo err1
    
    W_FILTRO = "M_F_COD = " & frm_Alt_Fic_Mensal_VIS.TXT_FUNC.BoundText & ""
    
    txt_Mes1 = frm_Alt_Fic_Mensal_VIS.ADOREG.Recordset.Fields("M_Mes")
    txt_Ano1 = frm_Alt_Fic_Mensal_VIS.ADOREG.Recordset.Fields("M_Ano")
    
    txt_Mes2 = frm_Alt_Fic_Mensal_VIS.ADOREG.Recordset.Fields("M_Mes")
    txt_Ano2 = frm_Alt_Fic_Mensal_VIS.ADOREG.Recordset.Fields("M_Ano")
    txt_Mes2 = CDbl(txt_Mes2) - 1
    If txt_Mes2 = 0 Then
        txt_Mes2 = 12
        txt_Ano2 = CDbl(txt_Ano2) - 1
    End If
    de.rsTAB_FICHA_MENS.Requery
    bt_OK1_Click
    de.rsTAB_FICHA_MENS.Requery
    bt_OK1_Click
    bt_OK2_Click
     
     
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


'*** Caption no navegador ***



Private Sub adoReg_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1
    
 If w_Hab = True And Not ADOREG.Recordset.EOF Then
    grid_Conta.Visible = True
    
    ADOREG.Caption = "REGISTRO : " & ADOREG.Recordset.AbsolutePosition & " / " & ADOREG.Recordset.RecordCount & IIf(W_LD_FILTRO = True, " (FILTRADO)", "")
         
    Set adoConta.Recordset = de.cnc.Execute("SELECT TAB_DESC_CALC.C_CODIGO, TAB_DESC_CALC.C_N_FICHA, TAB_DESC_CALC.C_DT, TAB_TP_CONTA.TP_DESC, TAB_DESC_CALC.C_TP_OP, TAB_DESC_CALC.C_VALOR, TAB_DESC_CALC.C_VISTO, TAB_DESC_CALC.C_DESC FROM TAB_DESC_CALC, TAB_TP_CONTA WHERE TAB_DESC_CALC.C_TP_CONTA = TAB_TP_CONTA.TP_COD AND (TAB_DESC_CALC.C_N_FICHA = " & ADOREG.Recordset.Fields("M_Nficha") & ") ORDER BY TAB_TP_CONTA.TP_DESC, TAB_DESC_CALC.C_TP_OP DESC").Clone
    adoConta.Refresh
   ' Pause 0.3
    Total
    TXT_TOTAL.Refresh
    grid_Conta.Visible = True

Else
    grid_Conta.Visible = False
    TXT_MENOS = ""
    TXT_MAIS = ""
    TXT_TOTAL = ""

End If

sair:
    Exit Sub
err1:
    If Not Err.Number = -2147217885 Then MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


'** Barra de Ferramenta ***
Private Sub BarraF_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
        Case "fechar": Fechar
        Case "editar": Editar
        Case "imprimir": Imprimir
        
    End Select
End Sub


'*** Rotinas ***


Sub Imprimir()
 
On Error GoTo err1
    
    FRM_IMP_F.TXT_LOGO = TXT_CAMPOS(0)
    FRM_IMP_F.TXT_MES = Format(ADOREG.Recordset.Fields("DATA"), "MM")
    FRM_IMP_F.TXT_ANO = TXT_ANO
    FRM_IMP_F.dbNome = TXT_CAMPOS(1)

    FRM_IMP_F.Show 1
    
    If FRM_IMP_F.txt_State = "F" Then
       MsgBox "Impressão Cancelada!", vbCritical
    Else
        If de.rscmdRelFichaMensal.State = 1 Then de.rscmdRelFichaMensal.Close
        If de.rscmdRelFichaMensal_TRIPA.State = 1 Then de.rscmdRelFichaMensal_TRIPA.Close
        
        de.cmdRelFichaMensal FRM_IMP_F.TXT_MES, FRM_IMP_F.TXT_ANO, FRM_IMP_F.dbNome, FRM_IMP_F.TXT_LOGO
        de.cmdRelFichaMensal_TRIPA FRM_IMP_F.TXT_MES, FRM_IMP_F.TXT_ANO, FRM_IMP_F.dbNome & "%", FRM_IMP_F.TXT_LOGO & "%"
        
            Set AdoItem1 = de.rscmdRelFichaMensal_TRIPA.Fields(6).Value
            Criar_RPT_TRIPA de.rscmdRelFichaMensal_TRIPA, AdoItem1
        
        
        rptFichaMensal.Show 1
'        rptFichaMensal_Tripa.PrintReport False, rptRangeAllPages
    End If
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Resume sair
    
End Sub







Private Sub Editar()

On Error GoTo Prox

If ADOREG.Recordset.Fields("BLOQ") = False Then
    
    W_POS = txtM_NFICHA
    
    If Not frm_Alt_Fic_Mensal.ADOREG.Recordset.State Then
        frm_Alt_Fic_Mensal.ADO_GRID.Recordset.Filter = "m_nficha = " & txtM_NFICHA
        frm_Alt_Fic_Mensal.ADOREG.Recordset.Filter = "m_nficha = " & txtM_NFICHA
    End If
    
Prox:
On Error GoTo err1
    frm_Alt_Fic_Mensal.Show 1
    
Else
    MsgBox "Não é possível alterar uma ficha do mês retrasado ao atual!", vbExclamation
End If

    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub


Private Sub Fechar()
On Error GoTo err1

    If de.rsTAB_FICHA_MENS.State = 1 Then
        de.rsTAB_DESC_CALC.Requery
        de.rsTAB_FICHA_MENS.Requery
    End If
    
    W_FILTRO = 0
    'ADOREG.Recordset.Filter = 0
    Unload Me
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
End Sub














Private Sub grid_Conta_DblClick()
    
If ADOREG.Recordset.Fields("BLOQ") = 0 Then
    

    frm_Alt_Desc_Calc.lb_form = "visualizar"
    frm_Alt_Desc_Calc.TXT_NFICHA_CAD = ADOREG.Recordset.Fields("M_Nficha")
    frm_Alt_Desc_Calc.LB_FUNC.Caption = TXT_CAMPOS(1).Text
    frm_Alt_Desc_Calc.Show 1

Else
    MsgBox "Não é possível alterar uma ficha anterior ao mês passado!", vbExclamation
End If

End Sub


Private Sub Grid_DblClick()
    Editar
End Sub

'--------- Ao Pressionar uma Tecla -----------

Private Sub grid_Conta_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 69 Then
        
        grid_Conta_DblClick
    ElseIf KeyCode = 13 Then
        KeyCode = 0
        TXT_CAMPOS(2).SetFocus
    Else
    
        Keys KeyCode, Shift
    End If
End Sub
Private Sub GRID_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub











Private Sub grid_conta2_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
    If KeyCode = 13 And Shift = 0 Then
        KeyCode = 0
        TXT_CAMPOS(10).SetFocus
    End If

End Sub

Private Sub mnuEdit_Click()
    Editar
End Sub

Private Sub mnuf1_Click()
    txt_Mes2.SetFocus

End Sub




Private Sub mnuF2_Click()
    txt_Ano2.SetFocus
End Sub

Private Sub mnuF3_Click()
    txt_Mes1.SetFocus
End Sub

Private Sub mnuF4_Click()
    txt_Ano1.SetFocus
End Sub

Private Sub mnuFechar_Click()
    Fechar
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift

End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift

End Sub

Private Sub txt_Ano1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
    Keys KeyCode, Shift

End Sub

Private Sub txt_Ano2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
    Keys KeyCode, Shift
End Sub

Private Sub TXT_CAMPOS_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
    If KeyCode = 13 And Shift = 0 Then
        KeyCode = 0
        If Not Index = 2 And Not Index = 3 And Not Index = 9 And Not Index = 10 Then SendKeys "{backspace}"
        SendKeys "{tab}"
    End If
End Sub





Private Sub txt_PAno_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub








Private Sub TXT_CAMPOS_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo err1
    If Index >= 3 And Index <= 6 Then
            If TXT_CAMPOS(2) = "" And TXT_CAMPOS(3) = "" Then
                de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_DT_ADM = NULL, M_DT_REG = NULL, M_FERIAS = '" & TXT_CAMPOS(4) & "', M_OBS = '" & TXT_CAMPOS(5) & "', M_ANOTACAO = '" & TXT_CAMPOS(6) & "' WHERE (M_NFICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & ")", w_reg
            ElseIf TXT_CAMPOS(2) = "" Then
                de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_DT_ADM = NULL, M_DT_REG = '" & TXT_CAMPOS(3) & "', M_FERIAS = '" & TXT_CAMPOS(4) & "', M_OBS = '" & TXT_CAMPOS(5) & "', M_ANOTACAO = '" & TXT_CAMPOS(6) & "' WHERE (M_NFICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & ")", w_reg
            ElseIf TXT_CAMPOS(3) = "" Then
                de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_DT_ADM = '" & TXT_CAMPOS(2) & "', M_DT_REG = NULL, M_FERIAS = '" & TXT_CAMPOS(4) & "', M_OBS = '" & TXT_CAMPOS(5) & "', M_ANOTACAO = '" & TXT_CAMPOS(6) & "' WHERE (M_NFICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & ")", w_reg
            Else
                de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_DT_ADM = '" & TXT_CAMPOS(2) & "', M_DT_REG = '" & TXT_CAMPOS(3) & "', M_FERIAS = '" & TXT_CAMPOS(4) & "', M_OBS = '" & TXT_CAMPOS(5) & "', M_ANOTACAO = '" & TXT_CAMPOS(6) & "' WHERE (M_NFICHA = " & ADOREG.Recordset.Fields("M_NFICHA") & ")", w_reg
            End If
    Else
            If TXT_CAMPOS(10) = "" And TXT_CAMPOS(9) = "" Then
                de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_DT_ADM = NULL, M_DT_REG = NULL, M_FERIAS = '" & TXT_CAMPOS(13) & "', M_OBS = '" & TXT_CAMPOS(12) & "', M_ANOTACAO = '" & TXT_CAMPOS(11) & "' WHERE (M_NFICHA = " & adoReg2.Recordset.Fields("M_NFICHA") & ")", w_reg
            ElseIf TXT_CAMPOS(10) = "" Then
                de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_DT_ADM = NULL, M_DT_REG = '" & TXT_CAMPOS(9) & "', M_FERIAS = '" & TXT_CAMPOS(13) & "', M_OBS = '" & TXT_CAMPOS(12) & "', M_ANOTACAO = '" & TXT_CAMPOS(11) & "' WHERE (M_NFICHA = " & adoReg2.Recordset.Fields("M_NFICHA") & ")", w_reg
            ElseIf TXT_CAMPOS(9) = "" Then
                de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_DT_ADM = '" & TXT_CAMPOS(10) & "', M_DT_REG = NULL, M_FERIAS = '" & TXT_CAMPOS(13) & "', M_OBS = '" & TXT_CAMPOS(12) & "', M_ANOTACAO = '" & TXT_CAMPOS(11) & "' WHERE (M_NFICHA = " & adoReg2.Recordset.Fields("M_NFICHA") & ")", w_reg
            Else
                de.cnc.Execute "UPDATE TAB_FICHA_MENS SET M_DT_ADM = '" & TXT_CAMPOS(10) & "', M_DT_REG = '" & TXT_CAMPOS(3) & "', M_FERIAS = '" & TXT_CAMPOS(13) & "', M_OBS = '" & TXT_CAMPOS(12) & "', M_ANOTACAO = '" & TXT_CAMPOS(11) & "' WHERE (M_NFICHA = " & adoReg2.Recordset.Fields("M_NFICHA") & ")", w_reg
            End If
    End If
    
    If w_reg = 0 Then MsgBox "Não foi salvo a alteração!", vbCritical


sair:
    Exit Sub
err1:
    MsgBox Err.Number & " :  " & Err.Description, vbCritical
End Sub

Private Sub txt_Mes1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
    Keys KeyCode, Shift

End Sub

Private Sub txt_Mes2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
    Keys KeyCode, Shift
End Sub

Private Sub TXT_TOTAL_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_MAIS_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub TXT_MENOS_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub TXT_ANO_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub
Private Sub txtM_MES_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub

Private Sub txtM_NFICHA_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys KeyCode, Shift
End Sub




' -------  Teclas de Atalhos --------

Sub Keys(ByRef KeyCode, Shift As Integer)
'*** Shift (4 = Alt) ***
If Shift = 4 Then
    Select Case KeyCode
    Case 70: ' "F"
            Fechar
    Case 69: ' "E"
            Editar
    End Select
ElseIf KeyCode = 116 And Shift = 0 Then
    KeyCode = 0
    Fechar
    
End If

End Sub





