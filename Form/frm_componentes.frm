VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{5C8CED40-8909-11D0-9483-00A0C91110ED}#1.0#0"; "MSDATREP.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BEC61919-E6C4-11D1-BE7D-C63815000000}#1.0#0"; "FLEXWIZ.OCX"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_componentes 
   Caption         =   "Form1"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13785
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   13785
   StartUpPosition =   3  'Windows Default
   Begin MSDataRepeaterLib.DataRepeater DataRepeater1 
      Height          =   1275
      Left            =   7680
      TabIndex        =   39
      Top             =   6240
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2249
      _StreamID       =   -1412567295
      _Version        =   393216
      Caption         =   "DataRepeater1"
      BeginProperty RepeatedControlName {21FC0FC0-1E5C-11D1-A327-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
      EndProperty
   End
   Begin ComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   41
      Top             =   900
      Width           =   13785
      _ExtentX        =   24315
      _ExtentY        =   741
      Appearance      =   1
      _Version        =   327682
   End
   Begin MSFlexGridWizard.SubWizard SubWizard1 
      Height          =   4455
      Left            =   480
      TabIndex        =   47
      Top             =   5760
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7858
   End
   Begin ComctlLib.Slider Slider2 
      Height          =   675
      Left            =   8280
      TabIndex        =   46
      Top             =   6600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1191
      _Version        =   327682
   End
   Begin ComctlLib.ListView ListView3 
      Height          =   735
      Left            =   5640
      TabIndex        =   45
      Top             =   3480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ListView ListView2 
      Height          =   735
      Left            =   7560
      TabIndex        =   44
      Top             =   6840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.TreeView TreeView2 
      Height          =   735
      Left            =   4200
      TabIndex        =   43
      Top             =   7080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   1215
      Left            =   0
      TabIndex        =   42
      Top             =   7755
      Width           =   13785
      _ExtentX        =   24315
      _ExtentY        =   2143
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TabStrip TabStrip2 
      Height          =   735
      Left            =   9600
      TabIndex        =   40
      Top             =   5400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   315
      Left            =   7920
      TabIndex        =   38
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBList DBList1 
      Height          =   645
      Left            =   8040
      TabIndex        =   37
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1138
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   735
      Left            =   7440
      TabIndex        =   36
      Top             =   5760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   3240
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
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
      Caption         =   "Adodc1"
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
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   810
      Left            =   3240
      TabIndex        =   35
      Top             =   7560
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1429
      _CBWidth        =   3615
      _CBHeight       =   810
      _Version        =   "6.7.9782"
      MinHeight1      =   360
      Width1          =   2880
      NewRow1         =   0   'False
      MinHeight2      =   360
      Width2          =   1440
      NewRow2         =   -1  'True
      MinHeight3      =   360
      Width3          =   1440
      NewRow3         =   0   'False
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   8160
      TabIndex        =   34
      Top             =   6720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   645
      Left            =   7920
      TabIndex        =   33
      Top             =   6600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1138
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   735
      Left            =   7560
      TabIndex        =   32
      Top             =   6360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   735
      Left            =   5640
      TabIndex        =   31
      Top             =   5280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   8520
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   4200
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   7200
      Top             =   6360
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   6720
      TabIndex        =   30
      Top             =   6360
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   3960
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   2160
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3120
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   735
      Left            =   2880
      OleObjectBlob   =   "frm_componentes.frx":0000
      TabIndex        =   29
      Top             =   3840
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   735
      Left            =   7920
      TabIndex        =   28
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frm_componentes.frx":2356
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6000
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
      Height          =   255
      Left            =   7200
      TabIndex        =   27
      Top             =   6480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1572865
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   735
      Left            =   9600
      TabIndex        =   26
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      Format          =   83820545
      CurrentDate     =   38441
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   3600
      TabIndex        =   25
      Top             =   5880
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   83820545
      CurrentDate     =   38441
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   735
      Left            =   3360
      TabIndex        =   24
      Top             =   4440
      Width           =   255
      _ExtentX        =   423
      _ExtentY        =   1296
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   495
      Left            =   7800
      TabIndex        =   23
      Top             =   7320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      _Version        =   393216
      FullWidth       =   73
      FullHeight      =   33
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   3360
      TabIndex        =   22
      Top             =   4680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   675
      Left            =   7200
      TabIndex        =   21
      Top             =   6840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1111
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8640
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   735
      Left            =   7920
      TabIndex        =   20
      Top             =   6240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   735
      Left            =   8040
      TabIndex        =   19
      Top             =   6360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   18
      Top             =   8970
      Width           =   13785
      _ExtentX        =   24315
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   13785
      _ExtentX        =   24315
      _ExtentY        =   1588
      ButtonWidth     =   1455
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   735
      Left            =   8160
      TabIndex        =   16
      Top             =   6120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   735
      Left            =   3720
      TabIndex        =   15
      Top             =   2760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   735
      Left            =   4800
      TabIndex        =   14
      Top             =   3600
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   1296
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frm_componentes.frx":23DC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frm_componentes.frx":23F8
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin MSRDC.MSRDC MSRDC1 
      Height          =   735
      Left            =   7680
      Top             =   5400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      Options         =   0
      CursorDriver    =   0
      BOFAction       =   0
      EOFAction       =   0
      RecordsetType   =   1
      LockType        =   3
      QueryType       =   0
      Prompt          =   3
      Appearance      =   1
      QueryTimeout    =   30
      RowsetSize      =   100
      LoginTimeout    =   15
      KeysetSize      =   0
      MaxRows         =   0
      ErrorThreshold  =   -1
      BatchSize       =   15
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      ReadOnly        =   0   'False
      Appearance      =   -1  'True
      DataSourceName  =   ""
      RecordSource    =   ""
      UserName        =   ""
      Password        =   ""
      Connect         =   ""
      LogMessages     =   ""
      Caption         =   "MSRDC1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Width           =   1140
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   9480
      TabIndex        =   12
      Top             =   5760
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   2880
      TabIndex        =   11
      Top             =   3960
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2520
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   4080
      Top             =   7320
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   3840
      TabIndex        =   9
      Top             =   4080
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   7080
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   4200
      TabIndex        =   7
      Top             =   7560
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5040
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   8400
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   3240
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
   Begin Skin_Button.ctr_Button cmdSalarioGerente 
      Height          =   285
      Left            =   1320
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "Salários Gerentes"
      Top             =   1680
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   503
      BTYPE           =   2
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
      BCOL            =   14737632
      BCOLO           =   12632319
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_componentes.frx":2414
      PICN            =   "frm_componentes.frx":2430
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   1080
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   8454143
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   8454143
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":3712
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":53474
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":A31D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":F2F38
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":142C9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":1929FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":1E275E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":2324C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":282222
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":2D1F84
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":321CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":371A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":3C17AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":41150C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":46126E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":4B0FD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":500D32
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":550A94
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":5A07F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":5F0558
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":6402BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":69001C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":6DFD7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":72FAE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_componentes.frx":77F842
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   4560
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.OLE OLE1 
      Height          =   1215
      Left            =   4560
      TabIndex        =   13
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3840
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   4080
      X2              =   5280
      Y1              =   7440
      Y2              =   7920
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   3720
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   7320
      Width           =   1215
   End
End
Attribute VB_Name = "frm_componentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
