VERSION 5.00
Begin VB.Form FrmSplash 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   105
      Picture         =   "FrmSplash.frx":0000
      ScaleHeight     =   3705
      ScaleWidth      =   5400
      TabIndex        =   7
      Top             =   810
      Width           =   5430
   End
   Begin SIJOSoftInvoiceManager.SIJOButton SCmdInvoice 
      Height          =   360
      Left            =   5970
      TabIndex        =   2
      Top             =   1155
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   635
      BTYPE           =   5
      TX              =   "Invoice"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16777215
      FCOL            =   16777215
      FCOLO           =   16761024
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSplash.frx":4124A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SIJOSoftInvoiceManager.SIJOButton SCmdStock 
      Height          =   360
      Left            =   5970
      TabIndex        =   3
      Top             =   1815
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   635
      BTYPE           =   5
      TX              =   "Stock"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16777215
      FCOL            =   16777215
      FCOLO           =   16761024
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSplash.frx":41266
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SIJOSoftInvoiceManager.SIJOButton SCmdSales 
      Height          =   360
      Left            =   5970
      TabIndex        =   4
      Top             =   2475
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   635
      BTYPE           =   5
      TX              =   "Sales Ratio"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16777215
      FCOL            =   16777215
      FCOLO           =   16761024
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSplash.frx":41282
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SIJOSoftInvoiceManager.SIJOButton SCmdAbout 
      Height          =   360
      Left            =   5970
      TabIndex        =   5
      Top             =   3135
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   635
      BTYPE           =   5
      TX              =   "About"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16777215
      FCOL            =   16777215
      FCOLO           =   16761024
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSplash.frx":4129E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SIJOSoftInvoiceManager.SIJOButton SCmdExit 
      Height          =   360
      Left            =   5970
      TabIndex        =   6
      Top             =   3795
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   635
      BTYPE           =   5
      TX              =   "Exit"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16777215
      FCOL            =   16777215
      FCOLO           =   16761024
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSplash.frx":412BA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7830
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   5550
      TabIndex        =   1
      Top             =   810
      Width           =   2160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SIJO Soft Invoice Manager"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   7605
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
MkDir App.path & "\Data.{9D6EAA4F-27B2-4407-AC72-4BBD2FCB6ED1}"
End Sub

Private Sub SCmdAbout_Click()
    MsgBox "Designed for SOFINE by SIJO Soft Corp.", vbExclamation, "SIJO Soft"

End Sub

Private Sub SCmdExit_Click()
    End
End Sub

Private Sub SCmdInvoice_Click()
    FrmMain.Show
    Unload Me
End Sub

Private Sub SCmdSales_Click()
    MsgBox "Show day to dat customer name and purchase details", vbInformation
End Sub

Private Sub SCmdStock_Click()
    FrmStock.Show
    Unload Me
End Sub


