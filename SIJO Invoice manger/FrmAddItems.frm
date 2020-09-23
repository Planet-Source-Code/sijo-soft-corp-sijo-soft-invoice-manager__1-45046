VERSION 5.00
Begin VB.Form FrmAddItems 
   BackColor       =   &H00D9E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Product"
   ClientHeight    =   5475
   ClientLeft      =   3180
   ClientTop       =   1665
   ClientWidth     =   5535
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4380
      MaxLength       =   4
      TabIndex        =   2
      Text            =   "1"
      Top             =   360
      Width           =   915
   End
   Begin VB.PictureBox PicList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   120
      ScaleHeight     =   0
      ScaleWidth      =   5145
      TabIndex        =   19
      Top             =   1230
      Visible         =   0   'False
      Width           =   5175
      Begin SIJOSoftInvoiceManager.SIJOButton SCMDDelete 
         Height          =   255
         Left            =   75
         TabIndex        =   27
         Top             =   3585
         Visible         =   0   'False
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   450
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
         BCOL            =   13160664
         BCOLO           =   13160664
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmAddItems.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SIJOSoftInvoiceManager.SIJOButton SCMDAdd 
         Height          =   255
         Left            =   3930
         TabIndex        =   26
         Top             =   3585
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   450
         BTYPE           =   5
         TX              =   "Add"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160664
         BCOLO           =   13160664
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmAddItems.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ListBox LstCurentMItms 
         Appearance      =   0  'Flat
         Height          =   3345
         Left            =   -15
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   195
         Width           =   5175
      End
      Begin VB.Label LblPicList 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Currently Available main Items"
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
         Height          =   240
         Left            =   0
         TabIndex        =   21
         Top             =   -15
         Width           =   5145
      End
   End
   Begin SIJOSoftInvoiceManager.SIJOButton SCmdSave 
      Height          =   360
      Left            =   3960
      TabIndex        =   10
      Top             =   4785
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
      BTYPE           =   5
      TX              =   "Save"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160664
      BCOLO           =   13160664
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmAddItems.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SIJOSoftInvoiceManager.SIJOButton SCmdCancel 
      Height          =   360
      Left            =   135
      TabIndex        =   11
      Top             =   4785
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
      BTYPE           =   5
      TX              =   "Cancel"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160664
      BCOLO           =   13160664
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmAddItems.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtPrchsDt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   135
      TabIndex        =   9
      Top             =   3990
      Width           =   5160
   End
   Begin VB.TextBox TxtVenderNme 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   135
      TabIndex        =   8
      Top             =   3420
      Width           =   5160
   End
   Begin VB.TextBox TxtRRPs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3825
      MaxLength       =   2
      TabIndex        =   7
      Text            =   "00"
      Top             =   2805
      Width           =   1470
   End
   Begin VB.TextBox TxtRRRs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   135
      TabIndex        =   6
      Top             =   2805
      Width           =   3705
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   480
      TabIndex        =   23
      Text            =   "Measurement"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1800
      TabIndex        =   22
      Text            =   "Measurement"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox TxtWRPs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3825
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "00"
      Top             =   1890
      Width           =   1470
   End
   Begin VB.TextBox TxtWRRs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   150
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1890
      Width           =   3705
   End
   Begin VB.TextBox TxtMItmNme 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5175
   End
   Begin VB.TextBox TxtPrdctNme 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4275
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00D9E9EC&
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity :"
      Height          =   195
      Left            =   4455
      TabIndex        =   28
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Date :"
      Height          =   195
      Left            =   135
      TabIndex        =   25
      Top             =   3780
      Width           =   1155
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vender Name :"
      Height          =   195
      Left            =   135
      TabIndex        =   24
      Top             =   3180
      Width           =   1065
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ps."
      Height          =   195
      Left            =   4425
      TabIndex        =   14
      Top             =   2550
      Width           =   225
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
      Height          =   195
      Left            =   1110
      TabIndex        =   15
      Top             =   2550
      Width           =   240
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Retail Rate :"
      Height          =   195
      Left            =   135
      TabIndex        =   16
      Top             =   2235
      Width           =   885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ps."
      Height          =   195
      Left            =   4425
      TabIndex        =   13
      Top             =   1620
      Width           =   225
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
      Height          =   195
      Left            =   1110
      TabIndex        =   17
      Top             =   1620
      Width           =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WholeSale Rate :"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   1335
      Width           =   1260
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Under,"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product name :"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "FrmAddItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ItemCode As String

Private Sub Form_Load()
    PicList.Height = 3915
End Sub

Private Sub RRRs_KeyPress(KeyAscii As Integer)

End Sub

Private Sub LstCurentMItms_Click()
'IF DELETE KEY IS WORKING MOVE IT TO DblClick
    TxtMItmNme.Text = LstCurentMItms.Text
    PicList.Visible = False
    TxtWRRs.SetFocus
End Sub

Private Sub LstCurentMItms_DblClick()
    'DELETE KEY IS NOT WORKING
    'TxtMItmNme.Text = LstCurentMItms.Text
    'PicList.Visible = False
    'TxtWRRs.SetFocus
End Sub

Private Sub LstCurentMItms_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtMItmNme.Text = LstCurentMItms.Text
    PicList.Visible = False
    TxtWRRs.SetFocus
End If
End Sub

Private Sub SCMDAdd_Click()
    Dim SS
    SS = InputBox("Enter a new name", "Add New Product Group", , 3270, 3050)
    If SS = "" Then
        Set SS = Nothing
    Else: LstCurentMItms.AddItem SS
    TxtMItmNme.Text = SS
    PicList.Visible = False
    Call GiveItemCode(TxtPrdctNme.Text)
    Dim Sijo As New SIJOINI
    With Sijo
        .path = App.path & "\Data.{9D6EAA4F-27B2-4407-AC72-4BBD2FCB6ED1}\MNDT.Dat"
        .Section = SS
        .Key = TxtPrdctNme.Text
        .Value = ItemCode
    End With
    End If

End Sub

Private Sub SCmdCancel_Click()
Dim SIJOMsg
SIJOMsg = MsgBox("Cancel?", vbYesNo + vbQuestion, App.Title)
If SIJOMsg = vbYes Then
    Unload Me
Else
    Exit Sub
End If
End Sub

Private Sub SCMDDelet_Click()

End Sub

Private Sub TxtMItmName_Change()

End Sub

Private Sub TxtMItmName_GotFocus()
    
End Sub

Private Sub SCmdSave_Click()
If TxtPrdctNme.Text = "" Then
    MsgBox "Please Fill all the fields.", vbInformation, App.Title
    Exit Sub
ElseIf TxtQty.Text = "" Then
    MsgBox "Please Fill all the fields.", vbInformation, App.Title
    Exit Sub
ElseIf TxtMItmNme.Text = "" Then
    MsgBox "Please Fill all the fields.", vbInformation, App.Title
    Exit Sub
ElseIf TxtWRRs.Text = "" Then
    MsgBox "Please Fill all the fields.", vbInformation, App.Title
    Exit Sub
ElseIf TxtRRRs.Text = "" Then
    MsgBox "Please Fill all the fields.", vbInformation, App.Title
    Exit Sub
ElseIf TxtVenderNme.Text = "" Then
    MsgBox "Please Fill all the fields.", vbInformation, App.Title
    Exit Sub
ElseIf TxtPrchsDt.Text = "" Then
    MsgBox "Please Fill all the fields.", vbInformation, App.Title
    Exit Sub
End If
Dim SIJOMsg
SIJOMsg = MsgBox("Save?", vbYesNo + vbQuestion, App.Title)
If SIJOMsg = vbYes Then
    Dim Sijo As New SIJOINI
        With Sijo
            .path = App.path & "\Data.{9D6EAA4F-27B2-4407-AC72-4BBD2FCB6ED1}\MNDT.Dat"
            .Section = TxtMItmNme.Text
            Call GiveItemCode(TxtPrdctNme.Text)
            .Key = TxtPrdctNme.Text
            .Value = ItemCode
            '======================================
            .path = App.path & "\Data.{9D6EAA4F-27B2-4407-AC72-4BBD2FCB6ED1}\" & ItemCode & ".Dat"
            .Section = TxtPrdctNme.Text
            .Key = "Qty :"
            .Value = TxtQty.Text
            .Key = "WR :"
            .Value = TxtWRRs.Text & "." & TxtWRPs.Text
            '.Key = "WRW :"
            '.Value = "WholesaleRate in words"
            .Key = "RR :"
            .Value = TxtRRRs.Text & "." & TxtRRPs.Text
            '.Key = "RRW :"
            '.Value = "Retail Rate in words"
            .Key = "TtlWR"
            .Value = Val(TxtWRRs.Text & "." & TxtWRPs.Text) * TxtQty.Text
            .Key = "TtlRR"
            .Value = Val(TxtRRRs.Text & "." & TxtRRPs.Text) * TxtQty.Text
            .Key = "VndrNme :"
            .Value = TxtVenderNme.Text
            .Key = "purchase Date :"
            .Value = TxtPrchsDt.Text
        End With
Else
    Exit Sub
End If
LoadMainItems
ClearAll
End Sub


Private Sub SIJOButton2_Click()

End Sub

Private Sub TxtMItmNme_GotFocus()
    PicList.Visible = True
    LstCurentMItms.Clear
    Dim Sijo As SIJOINI
    Dim strSections() As String
    Dim lonSectionCount As Long
    Dim lonCurrentSection As Long
 
        Set Sijo = New SIJOINI
 
        With Sijo
            .path = App.path & "\Data.{9D6EAA4F-27B2-4407-AC72-4BBD2FCB6ED1}\MNDT.Dat"
            .EnumerateAllSections strSections(), lonSectionCount
 
            For lonCurrentSection = 1 To lonSectionCount
                 LstCurentMItms.AddItem strSections(lonCurrentSection)
            Next lonCurrentSection
        End With
    LstCurentMItms.SetFocus

End Sub

Private Sub TxtMItmNme_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtMItmNme.Text = LstCurentMItms.Text
        PicList.Visible = False
    End If
End Sub

Private Sub WRRs_KeyPress(KeyAscii As Integer)

End Sub

Private Sub GiveItemCode(ProductName As String)
Dim GroupId, ProductId
Dim Sijo As New SIJOINI
    With Sijo
        .path = App.path & "\Data.{9D6EAA4F-27B2-4407-AC72-4BBD2FCB6ED1}\CDS.Dat"
        .Section = "Product Id"
        .Key = "Product Id :"
        If .Value = .path Then .Value = 1 Else: .Value = .Value + 1
        ProductId = .Value
        .Section = ProductName
        .Key = "Group :"
        .Value = TxtMItmNme.Text
        .Key = "Group Id :"
        If Asc(Left(TxtMItmNme.Text, 1)) < 90 Then
            GroupId = Left(TxtMItmNme.Text, 1)
        ElseIf Asc(Left(TxtMItmNme.Text, 1)) > 90 Then
            GroupId = Chr(Asc(Left(TxtMItmNme.Text, 1)) - 32)
        End If
        If Asc(Right(TxtMItmNme.Text, 1)) < 90 Then
            GroupId = GroupId & Right(TxtMItmNme.Text, 1)
        ElseIf Asc(Right(TxtMItmNme.Text, 1)) > 90 Then
            GroupId = GroupId & Chr(Asc(Right(TxtMItmNme.Text, 1)) - 32)
        End If
        .Value = GroupId & ProductId
        ItemCode = .Value
    End With
End Sub

Private Sub TxtPrchsDt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SCmdSave.SetFocus
End If
End Sub

Private Sub TxtPrdctNme_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then TxtQty.SetFocus
End Sub

Private Sub TxtQty_GotFocus()
TxtQty.Text = ""
End Sub

Private Sub TxtQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then TxtMItmNme.SetFocus
End Sub

Private Sub TxtQty_LostFocus()
If TxtQty.Text = "" Then TxtQty.Text = 1
End Sub

Private Sub TxtRRPs_GotFocus()
TxtRRPs.Text = ""
End Sub

Private Sub TxtRRPs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtVenderNme.SetFocus
End If
End Sub

Private Sub TxtRRPs_LostFocus()
If TxtRRPs.Text = "" Then TxtRRPs.Text = "00"
End Sub

Private Sub TxtRRRs_KeyPress(KeyAscii As Integer)
If KeyAscii = "46" Then
    KeyAscii = 0
    TxtRRPs.SetFocus
ElseIf KeyAscii = 13 Then
    TxtRRPs.SetFocus
End If
End Sub

Private Sub TxtVenderNme_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtPrchsDt.SetFocus
End If
End Sub

Private Sub TxtWRPs_GotFocus()
TxtWRPs.Text = ""
End Sub

Private Sub TxtWRPs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtRRRs.SetFocus
End If
End Sub

Private Sub TxtWRPs_LostFocus()
If TxtWRPs.Text = "" Then TxtWRPs.Text = "00"
End Sub

Private Sub TxtWRRs_KeyPress(KeyAscii As Integer)
If KeyAscii = "46" Then
    KeyAscii = 0
    TxtWRPs.SetFocus
ElseIf KeyAscii = 13 Then
    TxtWRPs.SetFocus
End If
End Sub

Private Sub LoadMainItems()
    FrmMain.lstINI.Clear
    Dim Sijo As SIJOINI
    Dim strSections() As String
    Dim lonSectionCount As Long
    Dim lonCurrentSection As Long
        Set Sijo = New SIJOINI
        With Sijo
            .path = App.path & "\Data.{9D6EAA4F-27B2-4407-AC72-4BBD2FCB6ED1}\MNDT.Dat"
            .EnumerateAllSections strSections(), lonSectionCount
            For lonCurrentSection = 1 To lonSectionCount
                 FrmMain.lstINI.AddItem strSections(lonCurrentSection)
            Next lonCurrentSection
        End With
End Sub

Private Sub ClearAll()
TxtPrdctNme.Text = ""
TxtQty.Text = ""
TxtMItmNme.Text = ""
TxtWRRs.Text = ""
TxtWRPs.Text = ""
TxtRRRs.Text = ""
TxtRRPs.Text = ""
TxtVenderNme.Text = ""
TxtPrchsDt.Text = ""
Unload Me
End Sub

