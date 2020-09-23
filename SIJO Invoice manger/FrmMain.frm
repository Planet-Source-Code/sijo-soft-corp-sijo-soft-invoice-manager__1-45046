VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "S I J O  S o f t    I n v o i c e   M a n a g e r ."
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.ListBox LstRRate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6660
      ItemData        =   "FrmMain.frx":08E1
      Left            =   10140
      List            =   "FrmMain.frx":08E3
      TabIndex        =   16
      Top             =   345
      Width           =   675
   End
   Begin VB.ListBox LstAmt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6660
      ItemData        =   "FrmMain.frx":08E5
      Left            =   10800
      List            =   "FrmMain.frx":08E7
      TabIndex        =   11
      Top             =   345
      Width           =   960
   End
   Begin VB.ListBox LstWRate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6660
      ItemData        =   "FrmMain.frx":08E9
      Left            =   9480
      List            =   "FrmMain.frx":08EB
      TabIndex        =   10
      Top             =   345
      Width           =   675
   End
   Begin SIJOSoftInvoiceManager.SIJOButton SCmdPrint 
      Height          =   330
      Left            =   10620
      TabIndex        =   7
      Top             =   7590
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   582
      BTYPE           =   5
      TX              =   "Print"
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
      MICON           =   "FrmMain.frx":08ED
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SIJOSoftInvoiceManager.SIJOButton SIJOButton1 
      Height          =   330
      Left            =   4935
      TabIndex        =   6
      Top             =   7590
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   582
      BTYPE           =   5
      TX              =   "Calculator"
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
      MICON           =   "FrmMain.frx":0909
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstTtlQty 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6660
      ItemData        =   "FrmMain.frx":0925
      Left            =   8760
      List            =   "FrmMain.frx":0927
      TabIndex        =   5
      Top             =   345
      Width           =   735
   End
   Begin VB.ListBox LstTtlnme 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6660
      ItemData        =   "FrmMain.frx":0929
      Left            =   4935
      List            =   "FrmMain.frx":092B
      TabIndex        =   4
      Top             =   345
      Width           =   3840
   End
   Begin SIJOSoftInvoiceManager.SIJOButton SCMDAddPrdct 
      Height          =   300
      Left            =   3165
      TabIndex        =   3
      Top             =   8385
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      BTYPE           =   5
      TX              =   "Add Product"
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
      MICON           =   "FrmMain.frx":092D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox PicLst 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3960
      Left            =   1320
      ScaleHeight     =   3930
      ScaleWidth      =   3105
      TabIndex        =   1
      Top             =   345
      Visible         =   0   'False
      Width           =   3135
      Begin VB.ListBox LstPrdct 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   -15
         TabIndex        =   2
         Top             =   210
         Width           =   3135
      End
      Begin VB.Label LblAProduct 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Available Products"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   3105
      End
   End
   Begin VB.ListBox lstINI 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8025
      Left            =   135
      TabIndex        =   0
      Top             =   345
      Width           =   4320
   End
   Begin VB.Label LblAmtChr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   4935
      TabIndex        =   22
      Top             =   7260
      Width           =   6825
   End
   Begin VB.Label LblTotalAmt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   10800
      TabIndex        =   21
      Top             =   6990
      Width           =   960
   End
   Begin VB.Label LblTotalRR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   10140
      TabIndex        =   20
      Top             =   6990
      Width           =   675
   End
   Begin VB.Label LblTotalWR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   9480
      TabIndex        =   19
      Top             =   6990
      Width           =   675
   End
   Begin VB.Label LblTotalQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   8760
      TabIndex        =   18
      Top             =   6990
      Width           =   735
   End
   Begin VB.Label LblRRate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RRate"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   10140
      TabIndex        =   17
      Top             =   135
      Width           =   675
   End
   Begin VB.Label LblAmt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Amount"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   10800
      TabIndex        =   15
      Top             =   135
      Width           =   960
   End
   Begin VB.Label LblWRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WRate"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   9480
      TabIndex        =   14
      Top             =   135
      Width           =   675
   End
   Begin VB.Label LblPProduct 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Purchased Product"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4935
      TabIndex        =   13
      Top             =   135
      Width           =   3840
   End
   Begin VB.Label LblQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quantity"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   8760
      TabIndex        =   12
      Top             =   135
      Width           =   735
   End
   Begin VB.Label LblAItem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Available Items"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   135
      TabIndex        =   8
      Top             =   135
      Width           =   4320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   4665
      X2              =   4665
      Y1              =   930
      Y2              =   7215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   4725
      X2              =   4725
      Y1              =   930
      Y2              =   7215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   2
      X1              =   4695
      X2              =   4695
      Y1              =   135
      Y2              =   8100
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LstiniIndex, LstPrdctIndex As Integer
Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
Call LoadMainItems
End Sub

Private Sub LoadMainItems()
    lstINI.Clear
    Dim Sijo As SIJOINI
    Dim strSections() As String
    Dim lonSectionCount As Long
    Dim lonCurrentSection As Long
        Set Sijo = New SIJOINI
        With Sijo
            .path = App.path & "\Data.{9D6EAA4F-27B2-4407-AC72-4BBD2FCB6ED1}\MNDT.Dat"
            .EnumerateAllSections strSections(), lonSectionCount
            For lonCurrentSection = 1 To lonSectionCount
                 lstINI.AddItem strSections(lonCurrentSection)
            Next lonCurrentSection
        End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmSplash.Show
End Sub

Private Sub LblTotalAmt_Change()
    If LblTotalAmt.Caption <> "" Then
        LblAmtChr.Caption = "Rs." & inttoword(LblTotalAmt.Caption)
    Else
        LblAmtChr.Caption = ""
    End If
End Sub

Private Sub LstAmt_Click()
    LstTtlnme.ListIndex = LstAmt.ListIndex
End Sub

Private Sub lstINI_Click()
If lstINI.ListIndex < 35 Then
    PicLst.Top = (lstINI.ListIndex + 3) * 120
    LoadProducts
End If
End Sub

Private Sub LoadProducts()
   lstprdct.Clear
   Dim Sijo As SIJOINI
   Dim strKeys() As String
   Dim lonKeyCount As Long
   Dim lonCurrentKey As Long
       Set Sijo = New SIJOINI
       With Sijo
           .path = App.path & "\Data.{9D6EAA4F-27B2-4407-AC72-4BBD2FCB6ED1}\MNDT.Dat"
           .Section = lstINI.Text
           .EnumerateCurrentSection strKeys(), lonKeyCount
           For lonCurrentKey = 1 To lonKeyCount
               .Key = strKeys(lonCurrentKey)
               lstprdct.AddItem .Key
           Next lonCurrentKey
       End With
End Sub

Private Sub lstINI_DblClick()
    LstiniIndex = lstINI.ListIndex
    'Call LoadMainItems
    KeyCode = 0
    PicLst.Visible = True
    lstprdct.SetFocus
    lstprdct.ListIndex = 0
End Sub

Private Sub lstINI_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
    lstINI_KeyPress (45)
    KeyCode = 0
ElseIf KeyCode = 40 Then
    lstINI_KeyPress (43)
    KeyCode = 0
ElseIf KeyCode = 39 Then
    LstiniIndex = lstINI.ListIndex
    'Call LoadMainItems
    KeyCode = 0
    PicLst.Visible = True
    lstprdct.SetFocus
    lstprdct.ListIndex = 0
End If
End Sub

Private Sub lstINI_KeyPress(KeyAscii As Integer)
If KeyAscii = 45 Then ' Minus key
    If lstINI.ListIndex > 0 Then
        lstINI.ListIndex = lstINI.ListIndex - 1
    Else
        lstINI.ListIndex = lstINI.ListCount - 1
    End If
ElseIf KeyAscii = 43 Then ' Plus key
    If lstINI.ListIndex < lstINI.ListCount - 1 Then
        lstINI.ListIndex = lstINI.ListIndex + 1
    Else
        lstINI.ListIndex = lstINI.ListIndex = 1
    End If
ElseIf KeyAscii = 13 Or KeyAscii = 42 Then
    LstiniIndex = lstINI.ListIndex
    'Call LoadMainItems
    KeyCode = 0
    PicLst.Visible = True
    lstprdct.SetFocus
    lstprdct.ListIndex = 0
End If
End Sub

Private Sub LstPrdct_DblClick()
    LstPrdct_KeyPress (13)
End Sub

Private Sub LstPrdct_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
    LstPrdct_KeyPress (45)
    KeyCode = 0
ElseIf KeyCode = 40 Then
    LstPrdct_KeyPress (43)
    KeyCode = 0
ElseIf KeyCode = 37 Then
    LstPrdctIndex = lstprdct.ListIndex
    'Call LoadProducts TO DISABLE SELECTION
    KeyCode = 0
    lstINI.ListIndex = LstiniIndex
    lstINI.SetFocus
    PicLst.Visible = False
End If
End Sub

Private Sub LstPrdct_KeyPress(KeyAscii As Integer)
If KeyAscii = 47 Then
    Call LstPrdct_KeyDown(37, 0)
ElseIf KeyAscii = 45 Then
    If lstprdct.ListIndex > 0 Then
        lstprdct.ListIndex = lstprdct.ListIndex - 1
    Else
        lstprdct.ListIndex = lstprdct.ListCount - 1
    End If
ElseIf KeyAscii = 43 Then
    If lstprdct.ListIndex < lstprdct.ListCount - 1 Then
        lstprdct.ListIndex = lstprdct.ListIndex + 1
    Else
        lstprdct.ListIndex = 0 'LstPrdct.ListIndex - 1
    End If
ElseIf KeyAscii = 13 Then
    prdctqty = InputBox("Quantity", "Enter Quantity", , 3270, 3050)
    If Not prdctqty = "" Then
        LstTtlnme.AddItem lstprdct.Text
        lstTtlQty.AddItem prdctqty
        LblTotalQty.Caption = Val(LblTotalQty.Caption) + Val(prdctqty)
        Dim Sijo As New SIJOINI
            With Sijo
                .path = App.path & "\Data.{9D6EAA4F-27B2-4407-AC72-4BBD2FCB6ED1}\MNDT.DAT"
                .Section = lstINI.Text
                .Key = lstprdct.Text
                Dim PCode
                PCode = .Value
                .path = App.path & "\Data.{9D6EAA4F-27B2-4407-AC72-4BBD2FCB6ED1}\" & PCode & ".DAT"
                .Section = lstprdct.Text
                .Key = "WR :"
                LstWRate.AddItem .Value
                LblTotalWR.Caption = Val(LblTotalWR.Caption) + Val(.Value)
                .Key = "RR :"
                LstRRate.AddItem .Value
                LblTotalRR.Caption = Val(LblTotalRR.Caption) + Val(.Value)
                LstAmt.AddItem .Value * prdctqty
                LblTotalAmt.Caption = Val(LblTotalAmt.Caption) + Val(.Value * prdctqty)
            End With
        
        PicLst.Visible = False
        lstINI.ListIndex = LstiniIndex
        lstINI.SetFocus
    Else
        Call LstPrdct_KeyDown(37, 0)
    End If
End If
End Sub

Private Sub LstRRate_Click()
    LstTtlnme.ListIndex = LstRRate.ListIndex
End Sub

Private Sub LstTtlnme_Click()
    lstTtlQty.ListIndex = LstTtlnme.ListIndex
    lstTtlQty.ListIndex = LstTtlnme.ListIndex
    LstWRate.ListIndex = LstTtlnme.ListIndex
    LstRRate.ListIndex = LstTtlnme.ListIndex
    LstAmt.ListIndex = LstTtlnme.ListIndex
End Sub

Private Sub lstTtlQty_Click()
    LstTtlnme.ListIndex = lstTtlQty.ListIndex
End Sub

Private Sub LstWRate_Click()
    LstTtlnme.ListIndex = LstWRate.ListIndex
End Sub

Private Sub SCMDAddPrdct_Click()
    FrmAddItems.Show , Me
End Sub

Private Sub SIJOButton2_Click()

End Sub

Private Sub SCmdPrint_Click()
With FrmBill
Dim Cnme
Cnme = InputBox("Enter Customer Name :", "Customer Name", , 3270, 3050)
If Cnme = "" Then
    Exit Sub
Else
    
    .TxtBill.Text = "Name : "
    .TxtBill.Text = .TxtBill.Text & Cnme
    Dim a
    Dim count
    a = 1
    count = 110 - Len(.TxtBill.Text)
    Do Until a = count
        .TxtBill.Text = .TxtBill.Text & " "
    a = a + 1
    Loop
    .TxtBill.Text = .TxtBill.Text & "Date : " & Date
    .Show , Me
    '.TxtBill.Text = .TxtBill.Text & LstTtlnme.Text
End If
End With
End Sub

Private Sub SIJOButton1_Click()
    Shell "explorer calc.exe"
End Sub
