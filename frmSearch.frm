VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form2"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10725
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   10725
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picSearchlist 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFDE91&
      ForeColor       =   &H80000008&
      Height          =   3660
      Left            =   705
      ScaleHeight     =   3630
      ScaleWidth      =   2535
      TabIndex        =   39
      Top             =   1260
      Visible         =   0   'False
      Width           =   2565
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1620
         TabIndex        =   46
         Top             =   75
         Width           =   660
      End
      Begin VB.CommandButton Command2 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2295
         TabIndex        =   45
         Top             =   75
         Width           =   225
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3255
         IntegralHeight  =   0   'False
         Left            =   -15
         Sorted          =   -1  'True
         TabIndex        =   41
         Top             =   390
         Width           =   2565
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keyword History"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   90
         Width           =   1395
      End
      Begin VB.Label Label18 
         BackColor       =   &H00B49800&
         Height          =   360
         Left            =   15
         TabIndex        =   43
         Top             =   15
         Width           =   2505
      End
   End
   Begin VB.PictureBox picKeywordHistoryShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3660
      Left            =   735
      ScaleHeight     =   3630
      ScaleWidth      =   2535
      TabIndex        =   44
      Top             =   1290
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.CommandButton Command1 
      Caption         =   ".."
      Height          =   270
      Left            =   3015
      TabIndex        =   40
      Top             =   960
      Width           =   225
   End
   Begin VB.PictureBox picExportMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFDE91&
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   5355
      ScaleHeight     =   1020
      ScaleWidth      =   2340
      TabIndex        =   33
      Top             =   1215
      Visible         =   0   'False
      Width           =   2370
      Begin VB.CommandButton cmdToHtml 
         Caption         =   "HTML"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -15
         TabIndex        =   34
         Top             =   705
         Width           =   2385
      End
      Begin VB.CommandButton cmdToText 
         Caption         =   "Text"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -15
         TabIndex        =   35
         Top             =   390
         Width           =   2385
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Export to"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   75
         Width           =   780
      End
      Begin VB.Label Label16 
         BackColor       =   &H00B49800&
         Height          =   360
         Left            =   15
         TabIndex        =   37
         Top             =   15
         Width           =   2310
      End
   End
   Begin VB.PictureBox picExportShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   5385
      ScaleHeight     =   1020
      ScaleWidth      =   2340
      TabIndex        =   38
      Top             =   1245
      Visible         =   0   'False
      Width           =   2370
   End
   Begin VB.PictureBox picOptMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFDE91&
      ForeColor       =   &H80000008&
      Height          =   3645
      Left            =   4470
      ScaleHeight     =   3615
      ScaleWidth      =   2340
      TabIndex        =   17
      Top             =   1215
      Visible         =   0   'False
      Width           =   2370
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   360
         ScaleHeight     =   300
         ScaleWidth      =   1530
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3030
         Width           =   1560
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   30
            ScaleHeight     =   240
            ScaleWidth      =   1470
            TabIndex        =   29
            Top             =   30
            Width           =   1470
            Begin VB.ComboBox cmbMax 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "frmSearch.frx":0442
               Left            =   -30
               List            =   "frmSearch.frx":0476
               TabIndex        =   30
               Text            =   "Combo1"
               Top             =   -30
               Width           =   1545
            End
         End
      End
      Begin VB.CheckBox chkFilename 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFDE91&
         Caption         =   "Filename"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   360
         TabIndex        =   23
         Top             =   555
         Width           =   1020
      End
      Begin VB.CheckBox chkAlbum 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFDE91&
         Caption         =   "ID3 Tag: Album"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   360
         TabIndex        =   22
         Top             =   840
         Width           =   1590
      End
      Begin VB.CheckBox chkArtist 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFDE91&
         Caption         =   "ID3 Tag: Artist"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   360
         TabIndex        =   21
         Top             =   1125
         Width           =   1590
      End
      Begin VB.CheckBox chkGenre 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFDE91&
         Caption         =   "ID3 Tag: Genre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   360
         TabIndex        =   20
         Top             =   1695
         Width           =   1485
      End
      Begin VB.CheckBox chkTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFDE91&
         Caption         =   "ID3 Tag: Title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   360
         TabIndex        =   19
         Top             =   1410
         Width           =   1380
      End
      Begin VB.CheckBox chkYear 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFDE91&
         Caption         =   "ID3 Tag: Year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   360
         TabIndex        =   18
         Top             =   1980
         Width           =   1485
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "# Records to Show"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   2580
         Width           =   1590
      End
      Begin VB.Label Label14 
         BackColor       =   &H00B49800&
         Height          =   360
         Left            =   15
         TabIndex        =   32
         Top             =   2520
         Width           =   2310
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fields to Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   90
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackColor       =   &H00B49800&
         Height          =   360
         Left            =   15
         TabIndex        =   26
         Top             =   15
         Width           =   2310
      End
   End
   Begin VB.PictureBox picOptShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3645
      Left            =   4500
      ScaleHeight     =   3645
      ScaleWidth      =   2370
      TabIndex        =   16
      Top             =   1245
      Visible         =   0   'False
      Width           =   2370
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Show Large Icons"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   705
      TabIndex        =   15
      Top             =   1305
      Width           =   1830
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   9270
      Top             =   5055
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   9150
      Top             =   8280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":04BF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9885
      Top             =   8310
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":0911
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7215
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   " Double-Click to Play a Song "
      Top             =   1800
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   12726
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList2"
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Volume"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Filename"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Album"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Artist"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Genre"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Title"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Year"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Comment"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Record"
         Object.Width           =   0
      EndProperty
      Picture         =   "frmSearch.frx":0D63
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3360
      TabIndex        =   1
      Top             =   930
      Width           =   885
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   705
      TabIndex        =   0
      Top             =   930
      Width           =   2565
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1530
      Left            =   390
      TabIndex        =   4
      Top             =   1785
      Width           =   99999
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8550
         TabIndex        =   13
         Top             =   915
         Width           =   1035
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   165
         TabIndex        =   5
         Top             =   540
         Visible         =   0   'False
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblCancelFlag 
         Caption         =   "0"
         Height          =   285
         Left            =   8535
         TabIndex        =   14
         Top             =   1245
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   165
         TabIndex        =   7
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Performing Search ..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   165
         TabIndex        =   6
         Top             =   285
         Width           =   4845
      End
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Export"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5370
      MouseIcon       =   "frmSearch.frx":2361
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   1020
      Width           =   765
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4485
      MouseIcon       =   "frmSearch.frx":266B
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   45
      Left            =   0
      TabIndex        =   12
      Top             =   1650
      Width           =   1.00005e5
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   705
      TabIndex        =   11
      Top             =   165
      Width           =   885
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   15
      Left            =   705
      TabIndex        =   10
      Top             =   540
      Width           =   1.00005e5
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmSearch.frx":2975
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   180
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find Word or Phrase"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   735
      TabIndex        =   2
      Top             =   705
      Width           =   1470
   End
   Begin VB.Image imgLogo 
      Height          =   1650
      Left            =   7155
      Picture         =   "frmSearch.frx":2DB7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3570
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Height          =   1695
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1.00005e5
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

    Select Case Check1.Value
    
    Case Checked
    ListView1.View = lvwIcon
    
    Case Unchecked
    ListView1.View = lvwReport
    
    End Select
    
End Sub

Private Sub chkAlbum_Click()

    Config.sAlbum = chkAlbum.Value
    SaveConfig
    
End Sub

Private Sub chkArtist_Click()

    Config.sArtist = chkArtist.Value
    SaveConfig
    
End Sub

Private Sub chkFilename_Click()

    Config.sFilename = chkFilename.Value
    SaveConfig
    
End Sub

Private Sub chkGenre_Click()

    Config.sGenre = chkGenre.Value
    SaveConfig
    
End Sub

Private Sub chkTitle_Click()

    Config.sTitle = chkTitle.Value
    SaveConfig
    
End Sub

Private Sub chkYear_Click()

    Config.sYear = chkYear.Value
    SaveConfig
    
End Sub

Private Sub cmbMax_Change()

    Config.MaxList = Val(cmbMax.Text)
    SaveConfig
    
End Sub

Private Sub cmbMax_Click()

    Config.MaxList = Val(cmbMax.Text)
    SaveConfig
    
End Sub

Private Sub cmbMax_KeyPress(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub

Private Sub cmdCancel_Click()

    lblCancelFlag.Caption = "1"

End Sub

Private Sub cmdSearch_Click()

    If Len(Trim$(Text4.Text)) = 0 Then
        MsgBox "Please enter a search word or phrase.", vbExclamation
        Exit Sub
    End If
    
    doAddKeyword
    
    Label4.Caption = ""
    
    MousePointer = 13
    ListView1.Visible = False
    ListView1.Sorted = False
    Refresh
    MaxVal% = cmbMax.Text
    
    With ProgressBar1
        .min = 0
        .Max = 1
        .Value = 0
        .Visible = True
    End With
    
    ListView1.ListItems.Clear
    
    Label4 = ""
    matches% = 0
    
    CatalogFileName = Trim$(Config.DBFileName)
    If InStr(1, CatalogFileName, "\") > 0 Then
        CatalogFileName = ExeDir & CatalogFileName
    End If
    
    ff = FreeFile
    Open CatalogFileName For Random As ff Len = Len(IndexFile)
    nRecs = LOF(ff) / Len(IndexFile)
    ProgressBar1.Max = nRecs + 1
    For vloop% = 1 To nRecs
    DoEvents
    Get ff, vloop, IndexFile
    
    ' BUILD SEARCH STRING
    sStr$ = ""
    With IndexFile
        If chkFilename.Value = Checked Then sStr = sStr & Trim$(.Filename)
        If chkAlbum.Value = Checked Then sStr = sStr & Trim$(.ID3.Album)
        If chkArtist.Value = Checked Then sStr = sStr & Trim$(.ID3.Artist)
        If chkTitle.Value = Checked Then sStr = sStr & Trim$(.ID3.Title)
        If chkGenre.Value = Checked Then sStr = sStr & Trim$(.ID3.Genre)
        If chkYear.Value = Checked Then sStr = sStr & Trim$(.ID3.Year)
    End With
    
    
    m = InStr(1, sStr, Text4.Text, vbTextCompare)
    If m > 0 And Trim$(sStr) <> "" Then
    
        With IndexFile
        tmpItem$ = "Vol " & Trim(.VolumeName) & "> " & Trim$(.Filename)
        End With
        
        '// ADD TO LISTVIEW
        With ListView1
        
            .ListItems.Add matches + 1, , "Vol " & Trim$(IndexFile.VolumeName), 1, 1
            .ListItems(matches + 1).SubItems(1) = Trim$(IndexFile.Filename)
            .ListItems(matches + 1).SubItems(2) = Trim$(IndexFile.ID3.Album)
            .ListItems(matches + 1).SubItems(3) = Trim$(IndexFile.ID3.Artist)
            .ListItems(matches + 1).SubItems(4) = Trim$(IndexFile.ID3.Genre)
            .ListItems(matches + 1).SubItems(5) = Trim$(IndexFile.ID3.Title)
            .ListItems(matches + 1).SubItems(6) = Trim$(IndexFile.ID3.Year)
            .ListItems(matches + 1).SubItems(7) = Trim$(IndexFile.ID3.Comment)
            .ListItems(matches + 1).SubItems(8) = Trim$(Str$(vloop))
            
        End With
        
        matches = matches + 1
        Label4.Caption = matches & " matches found (" & vloop & " records)"
        Label4.Refresh
        
    End If
    
    ProgressBar1.Value = vloop
    
    If matches >= MaxVal Then Exit For
    If lblCancelFlag.Caption = "1" Then
        lblCancelFlag.Caption = "0"
        Exit For
    End If
    
    DoEvents
    Next
    
    Close ff
        
    ProgressBar1.Visible = False
    Label4.Caption = "Found " & matches & " matches out of " & vloop & " records."

    Label7.Caption = "Search: [" & Text4.Text & "]    " & matches & " Matches"
    ListView1.Sorted = True
    ListView1.Visible = True
    MousePointer = 0
    
    If ListView1.ListItems.Count > 0 Then
        Me.cmdToText.Enabled = True
        Me.cmdToHtml.Enabled = True
    Else
        Me.cmdToText.Enabled = False
        Me.cmdToHtml.Enabled = False
    End If
    

End Sub

Private Sub cmdToHtml_Click()

    HideMenu
    
    '// SHOW DETAILS OF SELECTED RECORD
    If ListView1.ListItems.Count = 0 Then Exit Sub
        
    MousePointer = 11
    
    
    Dim hFile$, rFile$, HtmlString$
    hFile$ = ExeDir & "html-header-template.txt"
    rFile$ = ExeDir & "html-record-template.txt"
        
        
    '// PARSE TEMPLATES                                         //
    HtmlString = GetTextFromFile(hFile)
    HtmlString = Replace(HtmlString, "%search%", Label7.Caption)
    
    mbt$ = GetTextFromFile(rFile)
    For Y = 1 To ListView1.ListItems.Count
    
        '// GET RECORD DATA                                         //
        ff = FreeFile
        Open CatalogFileName For Random As ff Len = Len(IndexFile)
        '// Record number is in column 9
        Get ff, ListView1.ListItems(Y).SubItems(8), IndexFile
        Close ff
        
        With IndexFile
            ThisRecord$ = Replace(mbt$, "%volume%", Trim$(.VolumeName))
            ThisRecord$ = Replace(ThisRecord$, "%filename%", Trim$(.Filename))
        End With
        
        With IndexFile.ID3
            ThisRecord$ = Replace(ThisRecord$, "%album%", Trim$(.Album))
            ThisRecord$ = Replace(ThisRecord$, "%artist%", Trim$(.Artist))
            ThisRecord$ = Replace(ThisRecord$, "%comments%", Trim$(.Comment))
            ThisRecord$ = Replace(ThisRecord$, "%genre%", Trim$(.Genre))
            ThisRecord$ = Replace(ThisRecord$, "%title%", Trim$(.Title))
            ThisRecord$ = Replace(ThisRecord$, "%year%", Trim$(.Year))
        End With
        
        With IndexFile.Mp3Info
            ThisRecord$ = Replace(ThisRecord$, "%bitrate%", Trim$(.BitRate))
            ThisRecord$ = Replace(ThisRecord$, "%copyright%", Trim$(.Copy))
            ThisRecord$ = Replace(ThisRecord$, "%crc%", Trim$(.CRC))
            ThisRecord$ = Replace(ThisRecord$, "%emphasis%", Trim$(.Emphasis))
            ThisRecord$ = Replace(ThisRecord$, "%frequency%", Trim$(.FreqChannel))
            ThisRecord$ = Replace(ThisRecord$, "%layer%", Trim$(.Layer))
            ThisRecord$ = Replace(ThisRecord$, "%length%", Trim$(.Length))
            ThisRecord$ = Replace(ThisRecord$, "%original%", Trim$(.Original))
            ThisRecord$ = Replace(ThisRecord$, "%size%", Trim$(.Size))
        End With
        
        HtmlString = HtmlString & ThisRecord$
        ThisRecord$ = ""
        
    Next
    

    
    '// WRITE TO FILE
    ff = FreeFile
    Open "c:\dp-tmp.htm" For Output As ff
    Print #ff, HtmlString
    Close ff
    
    '// DISPLAY FILE IN BROWSER             //
    Dim Dummy As Long
    Dummy = ShellExecute(Me.hWnd, vbNullString, "c:\dp-tmp.htm", _
                         vbNullString, "c:\", 1)
    


    
    
    MousePointer = 0
    
End Sub

Private Sub cmdToText_Click()


    HideMenu
    
    '// SHOW DETAILS OF SELECTED RECORD
    If ListView1.ListItems.Count = 0 Then Exit Sub
        
    MousePointer = 11
    
        mbt$ = "Disk Pro" & vbCrLf _
             & Label7.Caption & vbCrLf _
             & "==================================================================" _
             & vbCrLf & vbCrLf
             
    For Y = 1 To ListView1.ListItems.Count
    
        mbt$ = mbt$ & "Volume: " & ListView1.ListItems(Y).Text & vbCrLf
        For X = 1 To 7
        mbt$ = mbt$ & ListView1.ColumnHeaders(X + 1).Text & ": " _
                & ListView1.ListItems(Y).SubItems(X) _
                & vbCrLf
        Next
        mbt$ = mbt$ & vbCrLf
        
    Next
    
    ff = FreeFile
    Open "c:\tmp.txt" For Output As ff
    Print #ff, mbt$
    Close ff
    MousePointer = 0
    
    
    Shell "notepad.exe c:\tmp.txt", vbNormalFocus
    
    
End Sub



Private Sub Command1_Click()

    ShowMenu 3
    
End Sub

Private Sub Command2_Click()

    HideMenu
    
End Sub

Private Sub Command3_Click()

    answer$ = MsgBox("Do you want to clear the keyword history?", _
                    vbYesNo + vbQuestion, "Are you sure?")
                    
    If answer <> vbYes Then Exit Sub
    
    '// CLEAR FILE  //
    kwfile$ = ExeDir & "\keywords.txt"
    ff = FreeFile
    Open kwfile For Random As ff
    Close ff
    Kill kwfile
    
    '// CLEAR LIST  //
    List1.Clear
    List1.SetFocus
    

End Sub

Private Sub Form_Load()

    cmbMax.Text = Trim(Config.MaxList)
    ListView1.View = Trim(Config.ListviewStyle)
    
    chkFilename.Value = Config.sFilename
    chkAlbum.Value = Config.sAlbum
    chkArtist.Value = Config.sArtist
    chkTitle.Value = Config.sTitle
    chkGenre.Value = Config.sGenre
    chkYear.Value = Config.sYear
    
    With ListView1
        For X = 1 To 9
            .ColumnHeaders(X).Width = Config.chWidth(X)
        Next
    End With
    
    MousePointer = 11
    doFillList
    MousePointer = 0
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    HideMenu
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Me.ProgressBar1.Width = ScaleWidth - 600
    ListView1.Width = ScaleWidth - 240
    ListView1.Height = ScaleHeight - ListView1.Top - 180
    Frame1.Move ListView1.Left, ListView1.Top, ListView1.Width, ListView1.Height
    cmdCancel.Left = ProgressBar1.Left + ProgressBar1.Width - cmdCancel.Width
    
    With imgLogo
        .Top = 15
        .Left = ScaleWidth - .Width - 15
    End With
    
End Sub

Private Sub HideMenu(Optional Index%)

    If Index Then
        
        Select Case Index
        
        Case 2
        Me.picOptMenu.Visible = False
        Me.picOptShadow.Visible = False
        
        Case 1
        Me.picExportMenu.Visible = False
        Me.picExportShadow.Visible = False
    
        End Select
        
    Else
        
        Me.picOptMenu.Visible = False
        Me.picOptShadow.Visible = False
        Me.picExportMenu.Visible = False
        Me.picExportShadow.Visible = False
        Me.picKeywordHistoryShadow.Visible = False
        Me.picSearchlist.Visible = False
    
    End If
    
End Sub

Private Sub ShowMenu(Index%)

    Select Case Index
    
    Case 2
    Me.picOptMenu.Visible = True
    Me.picOptShadow.Visible = True
    
    Case 1
    Me.picExportMenu.Visible = True
    Me.picExportShadow.Visible = True
    
    Case 3
    Me.picKeywordHistoryShadow.Visible = True
    Me.picSearchlist.Visible = True
    
    End Select
    
End Sub


Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ShowMenu 2
    HideMenu 1
    
    
End Sub



Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ShowMenu 1
    HideMenu 2
    
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    HideMenu
    
End Sub

Private Sub Label7_Change()

    Label5.Caption = Label7.Caption
    Caption = Label5.Caption
    
End Sub

Private Sub List1_Click()

    On Error Resume Next
    Me.Text4.Text = List1.List(List1.ListIndex)
    HideMenu
    cmdSearch_Click
    
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    ListView1.SortKey = ColumnHeader.Index - 1
    
    For X = 1 To 9
    ListView1.ColumnHeaders(X).Text = Trim$(Replace(ListView1.ColumnHeaders(X).Text, "*", ""))
    Next
    
    ListView1.ColumnHeaders(ColumnHeader.Index).Text = "* " & ListView1.ColumnHeaders(ColumnHeader.Index).Text
    
End Sub



Private Sub ListView1_DblClick()

    '// SHOW DETAILS OF SELECTED RECORD
    If ListView1.ListItems.Count = 0 Then Exit Sub
       
    MousePointer = 13
    
    mbt$ = "Volume: " & ListView1.SelectedItem.Text & vbCrLf
    For X = 1 To 8
    mbt$ = mbt$ & ListView1.ColumnHeaders(X + 1).Text & ": " _
                & ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(X) _
                & vbCrLf
    Next
    
    ff = FreeFile
    Open CatalogFileName For Random As ff Len = Len(IndexFile)
    '// Record number is in column 9
    Get ff, ListView1.SelectedItem.SubItems(8), IndexFile
    Close ff
        
    With IndexFile.Mp3Info
    mbt$ = mbt$ & "Size: " & .Size & vbCrLf _
                & "Length: " & .Length & vbCrLf _
                & "Layer: " & .Layer & vbCrLf _
                & "BitRate: " & .BitRate & vbCrLf _
                & "Frequency Channel: " & .FreqChannel & vbCrLf _
                & "CRC: " & .CRC & vbCrLf _
                & "Copyright: " & .Copy & vbCrLf _
                & "Emphasis: " & .Emphasis & vbCrLf _
                & "Original: " & .Original & vbCrLf
    End With
    
    PlayMP3 Trim$(Config.iDefaultPlayerDrive) _
          & ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1), _
          mbt$, ListView1.SelectedItem.Text
          
    MousePointer = 0
    
    
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    HideMenu
    
End Sub

Private Sub Text4_Change()

    cmdSearch.Default = True
    HideMenu
    
End Sub

Private Sub Timer1_Timer()

    Static OldSum
    Dim NewSum As Integer
    
    For X = 1 To 9
    NewSum% = NewSum% + ListView1.ColumnHeaders(X).Width
    Next
    
    If OldSum <> NewSum Then
        '// SAVE NEW COLUMN SIZES
        For X = 1 To 9
        Config.chWidth(X) = ListView1.ColumnHeaders(X).Width
        Next
        SaveConfig
    End If
    
    OldSum = NewSum
    
End Sub

Private Sub doAddKeyword()

    keyword$ = Text4.Text
    
    If List1.ListCount < 1 Then GoTo doAdd
    
    For X = 0 To List1.ListCount
        If LCase(List1.List(X)) = LCase$(keyword) Then
            Flag% = 1
            Exit For
        End If
    Next
    
    If Flag = 1 Then Exit Sub
    
    
    
doAdd:
    kwfile$ = ExeDir & "\keywords.txt"
    
    ff = FreeFile
    Open kwfile For Random As ff
    Close ff
    
    Open kwfile For Append As ff
    Print #ff, keyword
    Close ff
    
    Me.List1.AddItem keyword
    
End Sub

Private Sub doFillList()

    List1.Clear
    kwfile$ = ExeDir & "\keywords.txt"
    
    ff = FreeFile
    Open kwfile For Random As ff
    Close ff
    
    Open kwfile For Input As ff
    Do Until EOF(ff)
        Line Input #ff, a$
        List1.AddItem a$
    Loop
    Close ff
    
End Sub
