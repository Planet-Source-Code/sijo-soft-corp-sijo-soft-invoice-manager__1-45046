VERSION 5.00
Begin VB.Form FrmStock 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ListBox LstQty 
      Appearance      =   0  'Flat
      Height          =   8025
      Left            =   4410
      TabIndex        =   3
      Top             =   270
      Width           =   1080
   End
   Begin VB.ListBox Lsttemp 
      Appearance      =   0  'Flat
      Height          =   2955
      Left            =   6915
      TabIndex        =   2
      Top             =   1665
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox lstprdct 
      Appearance      =   0  'Flat
      Height          =   8025
      Left            =   105
      TabIndex        =   0
      Top             =   270
      Width           =   4320
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quantity"
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
      Height          =   225
      Left            =   4410
      TabIndex        =   4
      Top             =   60
      Width           =   1080
   End
   Begin VB.Label LblAItem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Available Items"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   105
      TabIndex        =   1
      Top             =   60
      Width           =   4320
   End
End
Attribute VB_Name = "FrmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Lsttemp.Clear
    Dim Sijo As SIJOINI
    Dim strSections() As String
    Dim lonSectionCount As Long
    Dim lonCurrentSection As Long
        Set Sijo = New SIJOINI
        With Sijo
            .path = App.path & "\Data.{9D6EAA4F-27B2-4407-AC72-4BBD2FCB6ED1}\MNDT.Dat"
            .EnumerateAllSections strSections(), lonSectionCount
            For lonCurrentSection = 1 To lonSectionCount
                 Lsttemp.AddItem strSections(lonCurrentSection)
            Next lonCurrentSection
        End With
    Dim s
    Do Until s = Lsttemp.ListCount
        Lsttemp.ListIndex = Lsttemp.ListIndex + 1
        s = s + 1
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmSplash.Show
End Sub

Private Sub Lsttemp_Click()
   Dim Sijo As SIJOINI
   Dim strKeys() As String
   Dim lonKeyCount As Long
   Dim lonCurrentKey As Long
       Set Sijo = New SIJOINI
       With Sijo
           .path = App.path & "\Data.{9D6EAA4F-27B2-4407-AC72-4BBD2FCB6ED1}\MNDT.Dat"
           .Section = Lsttemp.Text
           .EnumerateCurrentSection strKeys(), lonKeyCount
           For lonCurrentKey = 1 To lonKeyCount
               .Key = strKeys(lonCurrentKey)
               lstprdct.AddItem .Key
               Dim SKey
               SKey = .Key
                    Dim sam As New SIJOINI
                        sam.path = App.path & "\Data.{9D6EAA4F-27B2-4407-AC72-4BBD2FCB6ED1}\" & .Value & ".Dat"
                        sam.Section = SKey
                        sam.Key = "Qty :"
                        LstQty.AddItem sam.Value
           Next lonCurrentKey
       End With

End Sub
