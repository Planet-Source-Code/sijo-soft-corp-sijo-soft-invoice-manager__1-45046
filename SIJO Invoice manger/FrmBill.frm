VERSION 5.00
Begin VB.Form FrmBill 
   Caption         =   "Bill"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "FrmBill.frx":0000
      Top             =   660
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "FrmBill.frx":0065
      Top             =   75
      Width           =   6375
   End
   Begin VB.TextBox TxtBill 
      Appearance      =   0  'Flat
      Height          =   3360
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FrmBill.frx":008C
      Top             =   1335
      Width           =   6375
   End
End
Attribute VB_Name = "FrmBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
