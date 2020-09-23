VERSION 5.00
Object = "{01E1A639-5A2A-426E-A3E8-C0B0B1F968F1}#10.0#0"; "PGX.ocx"
Begin VB.Form frm1 
   Caption         =   "PureGold Button Control Now With the official ""Cool"" Button hover!"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox cap1 
      AutoSize        =   -1  'True
      Height          =   2490
      Left            =   4110
      Picture         =   "frm1.frx":0000
      ScaleHeight     =   2430
      ScaleWidth      =   4470
      TabIndex        =   22
      Top             =   780
      Width           =   4530
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   4320
      TabIndex        =   21
      Top             =   90
      Width           =   705
   End
   Begin PureGoldFlatPushButton.PGBC PGBC31 
      Height          =   375
      Left            =   90
      TabIndex        =   20
      Top             =   3000
      Width           =   6555
      _extentx        =   11562
      _extenty        =   661
      font            =   "frm1.frx":962C
      autoredraw      =   -1  'True
      fontbold        =   -1  'True
      fontname        =   "Poor Richard"
      fontsize        =   14.25
      scaleheight     =   25
      scalemode       =   3
      scalewidth      =   437
      caption         =   "PureGold Button Control (Exit)"
      borderstylec    =   5
   End
   Begin PureGoldFlatPushButton.PGBC PGBC25 
      Height          =   315
      Left            =   5490
      TabIndex        =   14
      Top             =   2190
      Width           =   1155
      _extentx        =   2037
      _extenty        =   556
      forecolor       =   255
      font            =   "frm1.frx":9658
      autoredraw      =   -1  'True
      fontname        =   "MS Sans Serif"
      fontsize        =   8.25
      scaleheight     =   21
      scalemode       =   3
      scalewidth      =   77
      caption         =   "ThickFlat"
      hover           =   -1  'True
      borderstylec    =   5
   End
   Begin PureGoldFlatPushButton.PGBC PGBC24 
      Height          =   315
      Left            =   4350
      TabIndex        =   13
      Top             =   2190
      Width           =   1095
      _extentx        =   1931
      _extenty        =   556
      forecolor       =   255
      font            =   "frm1.frx":9684
      autoredraw      =   -1  'True
      fontname        =   "MS Sans Serif"
      fontsize        =   8.25
      scaleheight     =   21
      scalemode       =   3
      scalewidth      =   73
      caption         =   "Outline"
      hover           =   -1  'True
      borderstylec    =   4
   End
   Begin PureGoldFlatPushButton.PGBC PGBC23 
      Height          =   315
      Left            =   2910
      TabIndex        =   12
      Top             =   2190
      Width           =   1395
      _extentx        =   2461
      _extenty        =   556
      forecolor       =   16711680
      font            =   "frm1.frx":96B0
      autoredraw      =   -1  'True
      fontname        =   "MS Sans Serif"
      fontsize        =   8.25
      scaleheight     =   21
      scalemode       =   3
      scalewidth      =   93
      caption         =   "Cplus"
      hover           =   -1  'True
      borderstylec    =   2
   End
   Begin PureGoldFlatPushButton.PGBC PGBC22 
      Height          =   315
      Left            =   1650
      TabIndex        =   11
      Top             =   2190
      Width           =   1215
      _extentx        =   2143
      _extenty        =   556
      forecolor       =   255
      font            =   "frm1.frx":96DC
      autoredraw      =   -1  'True
      fontname        =   "MS Sans Serif"
      fontsize        =   8.25
      scaleheight     =   21
      scalemode       =   3
      scalewidth      =   81
      caption         =   "DarkFlat"
      hover           =   -1  'True
      borderstylec    =   3
   End
   Begin PureGoldFlatPushButton.PGBC PGBC21 
      Height          =   315
      Left            =   90
      TabIndex        =   10
      Top             =   2190
      Width           =   1515
      _extentx        =   2672
      _extenty        =   556
      forecolor       =   255
      font            =   "frm1.frx":9708
      autoredraw      =   -1  'True
      fontname        =   "MS Sans Serif"
      fontsize        =   8.25
      scaleheight     =   21
      scalemode       =   3
      scalewidth      =   101
      caption         =   "Thin"
      hover           =   -1  'True
   End
   Begin PureGoldFlatPushButton.PGBC PGBC1 
      Height          =   345
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   1995
      _extentx        =   3519
      _extenty        =   609
      forecolor       =   16711680
      font            =   "frm1.frx":9734
      autoredraw      =   -1  'True
      fontbold        =   -1  'True
      fontname        =   "Times New Roman"
      fontsize        =   12
      scaleheight     =   23
      scalemode       =   3
      scalewidth      =   133
      caption         =   "Pure"
   End
   Begin PureGoldFlatPushButton.PGBC PGBC2 
      Height          =   345
      Left            =   30
      TabIndex        =   1
      Top             =   420
      Width           =   1995
      _extentx        =   3519
      _extenty        =   609
      forecolor       =   65535
      font            =   "frm1.frx":9764
      autoredraw      =   -1  'True
      fontbold        =   -1  'True
      fontname        =   "Times New Roman"
      fontsize        =   12
      scaleheight     =   23
      scalemode       =   3
      scalewidth      =   133
      caption         =   "Pure"
      borderstylec    =   3
   End
   Begin PureGoldFlatPushButton.PGBC PGBC3 
      Height          =   345
      Left            =   30
      TabIndex        =   2
      Top             =   810
      Width           =   1995
      _extentx        =   3519
      _extenty        =   609
      forecolor       =   16777088
      font            =   "frm1.frx":9794
      autoredraw      =   -1  'True
      fontbold        =   -1  'True
      fontname        =   "Times New Roman"
      fontsize        =   12
      scaleheight     =   23
      scalemode       =   3
      scalewidth      =   133
      caption         =   "Pure"
      borderstylec    =   5
   End
   Begin PureGoldFlatPushButton.PGBC PGBC4 
      Height          =   345
      Left            =   30
      TabIndex        =   3
      Top             =   1200
      Width           =   1995
      _extentx        =   3519
      _extenty        =   609
      forecolor       =   16576
      font            =   "frm1.frx":97C4
      autoredraw      =   -1  'True
      fontbold        =   -1  'True
      fontname        =   "Times New Roman"
      fontsize        =   12
      scaleheight     =   23
      scalemode       =   3
      scalewidth      =   133
      caption         =   "Pure"
      borderstylec    =   2
   End
   Begin PureGoldFlatPushButton.PGBC PGBC5 
      Height          =   345
      Left            =   30
      TabIndex        =   4
      Top             =   1590
      Width           =   1995
      _extentx        =   3519
      _extenty        =   609
      forecolor       =   16744576
      font            =   "frm1.frx":97F4
      autoredraw      =   -1  'True
      fontbold        =   -1  'True
      fontname        =   "Times New Roman"
      fontsize        =   12
      scaleheight     =   23
      scalemode       =   3
      scalewidth      =   133
      caption         =   "Pure"
      borderstylec    =   4
   End
   Begin PureGoldFlatPushButton.PGBC PGBC6 
      Height          =   345
      Left            =   2070
      TabIndex        =   5
      Top             =   30
      Width           =   1995
      _extentx        =   3519
      _extenty        =   609
      enabled         =   0   'False
      font            =   "frm1.frx":9824
      autoredraw      =   -1  'True
      fontbold        =   -1  'True
      fontname        =   "Times New Roman"
      fontsize        =   12
      scaleheight     =   23
      scalemode       =   3
      scalewidth      =   133
      caption         =   "Pure"
      focusrect       =   -1  'True
   End
   Begin PureGoldFlatPushButton.PGBC PGBC7 
      Height          =   345
      Left            =   2070
      TabIndex        =   6
      Top             =   420
      Width           =   1995
      _extentx        =   3519
      _extenty        =   609
      enabled         =   0   'False
      font            =   "frm1.frx":9854
      autoredraw      =   -1  'True
      fontbold        =   -1  'True
      fontname        =   "Times New Roman"
      fontsize        =   12
      scaleheight     =   23
      scalemode       =   3
      scalewidth      =   133
      caption         =   "Pure"
      focusrect       =   -1  'True
      borderstylec    =   3
   End
   Begin PureGoldFlatPushButton.PGBC PGBC8 
      Height          =   345
      Left            =   2070
      TabIndex        =   7
      Top             =   810
      Width           =   1995
      _extentx        =   3519
      _extenty        =   609
      enabled         =   0   'False
      font            =   "frm1.frx":9884
      autoredraw      =   -1  'True
      fontbold        =   -1  'True
      fontname        =   "Times New Roman"
      fontsize        =   12
      scaleheight     =   23
      scalemode       =   3
      scalewidth      =   133
      caption         =   "Pure"
      focusrect       =   -1  'True
      borderstylec    =   5
   End
   Begin PureGoldFlatPushButton.PGBC PGBC9 
      Height          =   345
      Left            =   2070
      TabIndex        =   8
      Top             =   1200
      Width           =   1995
      _extentx        =   3519
      _extenty        =   609
      enabled         =   0   'False
      font            =   "frm1.frx":98B4
      autoredraw      =   -1  'True
      fontbold        =   -1  'True
      fontname        =   "Times New Roman"
      fontsize        =   12
      scaleheight     =   23
      scalemode       =   3
      scalewidth      =   133
      caption         =   "Pure"
      focusrect       =   -1  'True
      borderstylec    =   2
   End
   Begin PureGoldFlatPushButton.PGBC PGBC10 
      Height          =   345
      Left            =   2070
      TabIndex        =   9
      Top             =   1590
      Width           =   1995
      _extentx        =   3519
      _extenty        =   609
      enabled         =   0   'False
      font            =   "frm1.frx":98E4
      autoredraw      =   -1  'True
      fontbold        =   -1  'True
      fontname        =   "Times New Roman"
      fontsize        =   12
      scaleheight     =   23
      scalemode       =   3
      scalewidth      =   133
      caption         =   "Pure"
      focusrect       =   -1  'True
      borderstylec    =   4
   End
   Begin PureGoldFlatPushButton.PGBC PGBC26 
      Height          =   315
      Left            =   5490
      TabIndex        =   15
      Top             =   2580
      Width           =   1155
      _extentx        =   2037
      _extenty        =   556
      forecolor       =   255
      font            =   "frm1.frx":9914
      autoredraw      =   -1  'True
      fontname        =   "Small Fonts"
      fontsize        =   6
      scaleheight     =   21
      scalemode       =   3
      scalewidth      =   77
      caption         =   "ThickFlat + Focus"
      focusrect       =   -1  'True
      hover           =   -1  'True
      borderstylec    =   5
   End
   Begin PureGoldFlatPushButton.PGBC PGBC27 
      Height          =   315
      Left            =   4350
      TabIndex        =   16
      Top             =   2580
      Width           =   1095
      _extentx        =   1931
      _extenty        =   556
      forecolor       =   255
      font            =   "frm1.frx":9940
      autoredraw      =   -1  'True
      fontname        =   "Small Fonts"
      fontsize        =   6.75
      scaleheight     =   21
      scalemode       =   3
      scalewidth      =   73
      caption         =   "OutLine + Focus"
      focusrect       =   -1  'True
      hover           =   -1  'True
      borderstylec    =   4
   End
   Begin PureGoldFlatPushButton.PGBC PGBC28 
      Height          =   315
      Left            =   2910
      TabIndex        =   17
      Top             =   2580
      Width           =   1395
      _extentx        =   2461
      _extenty        =   556
      forecolor       =   16711680
      font            =   "frm1.frx":996C
      autoredraw      =   -1  'True
      fontname        =   "MS Sans Serif"
      fontsize        =   8.25
      scaleheight     =   21
      scalemode       =   3
      scalewidth      =   93
      caption         =   "Cplus + Focus"
      focusrect       =   -1  'True
      hover           =   -1  'True
      borderstylec    =   2
   End
   Begin PureGoldFlatPushButton.PGBC PGBC29 
      Height          =   315
      Left            =   1650
      TabIndex        =   18
      Top             =   2580
      Width           =   1215
      _extentx        =   2143
      _extenty        =   556
      forecolor       =   255
      font            =   "frm1.frx":9998
      autoredraw      =   -1  'True
      fontname        =   "Small Fonts"
      fontsize        =   6.75
      scaleheight     =   21
      scalemode       =   3
      scalewidth      =   81
      caption         =   "DarkFlat + Focus"
      focusrect       =   -1  'True
      hover           =   -1  'True
      borderstylec    =   3
   End
   Begin PureGoldFlatPushButton.PGBC PGBC30 
      Height          =   315
      Left            =   90
      TabIndex        =   19
      Top             =   2580
      Width           =   1515
      _extentx        =   2672
      _extenty        =   556
      forecolor       =   255
      font            =   "frm1.frx":99C4
      autoredraw      =   -1  'True
      fontname        =   "MS Sans Serif"
      fontsize        =   8.25
      scaleheight     =   21
      scalemode       =   3
      scalewidth      =   101
      caption         =   "Thin + Focus"
      focusrect       =   -1  'True
      hover           =   -1  'True
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MaskPicture cap1
End Sub



Private Sub PGBC31_Click()
End
End Sub

