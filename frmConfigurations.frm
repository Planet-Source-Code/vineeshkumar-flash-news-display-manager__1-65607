VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmConfigurations 
   Caption         =   "Configure Outputs"
   ClientHeight    =   6375
   ClientLeft      =   8490
   ClientTop       =   930
   ClientWidth     =   6300
   Icon            =   "frmConfigurations.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   6300
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1080
      Top             =   5280
   End
   Begin VB.Frame frameScrollingNews 
      Caption         =   "Scrolling News"
      Height          =   5415
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox txtScrolltitle 
         Height          =   435
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtScrollDesc 
         Height          =   555
         Left            =   1080
         TabIndex        =   1
         Top             =   720
         Width           =   4935
      End
      Begin VB.CommandButton cmdScrollAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdScrollDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdScrollClear 
         Caption         =   "Cl&ear"
         Height          =   375
         Left            =   5040
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdScrollSet 
         Caption         =   "Set Scrolling News"
         Default         =   -1  'True
         Height          =   375
         Left            =   4440
         TabIndex        =   7
         Top             =   4920
         Width           =   1575
      End
      Begin MSComctlLib.ListView lstScroll 
         Height          =   3495
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   6165
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sl #"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Title"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date/Time"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblSeppic 
         AutoSize        =   -1  'True
         Caption         =   "cup.gif"
         Height          =   195
         Left            =   1920
         TabIndex        =   36
         Top             =   5040
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Seperator Image:"
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Image imgsep 
         Height          =   255
         Left            =   1560
         MouseIcon       =   "frmConfigurations.frx":000C
         MousePointer    =   99  'Custom
         Picture         =   "frmConfigurations.frx":0316
         Stretch         =   -1  'True
         Tag             =   "./cup.gif"
         ToolTipText     =   "click to browse"
         Top             =   5040
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Title"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Description"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   600
      Top             =   5160
   End
   Begin VB.OptionButton optTitle 
      Caption         =   "Title N.."
      Height          =   435
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   5880
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton optBN 
      Caption         =   "Break N.."
      Height          =   435
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   5880
      Width           =   855
   End
   Begin VB.OptionButton optScroll 
      Caption         =   "scroll"
      Height          =   435
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   5880
      Width           =   855
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   14
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   13
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   12
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Feedback"
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   5940
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Load Output"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   5940
      Width           =   1095
   End
   Begin VB.Frame frameBrknews 
      Caption         =   "Breaking News"
      Height          =   5175
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CheckBox chkEnableBN 
         Caption         =   "Enable Breaking News Display"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   4560
         Width           =   2535
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   375
         Left            =   4920
         TabIndex        =   37
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdsetBrknews 
         Caption         =   "Set Breaking News"
         Height          =   735
         Left            =   4200
         TabIndex        =   26
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtBrknews 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label lblBN 
         Caption         =   "Click the button to set the ""breaking news"""
         ForeColor       =   &H00404040&
         Height          =   1215
         Left            =   360
         TabIndex        =   27
         Top             =   1680
         Width           =   3495
      End
   End
   Begin VB.Frame frameTitleNews 
      Caption         =   "Title News"
      Height          =   5775
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6135
      Begin VB.ComboBox cmbEffects 
         Height          =   315
         ItemData        =   "frmConfigurations.frx":1A46
         Left            =   2400
         List            =   "frmConfigurations.frx":1A48
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   5310
         Width           =   2295
      End
      Begin VB.CheckBox chkTitle 
         Caption         =   "Enable Auto display"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   4920
         Width           =   1695
      End
      Begin VB.CommandButton cmdShootTitle 
         Caption         =   "&Shoot"
         Height          =   375
         Left            =   4800
         TabIndex        =   28
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CommandButton cmdClearTitle 
         Caption         =   "Cl&ear"
         Height          =   375
         Left            =   5040
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdDeleteTitle 
         Caption         =   "Delete"
         Height          =   375
         Left            =   4080
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin MSComctlLib.ListView lstTitles 
         Height          =   3495
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   6165
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sl #"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Title"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date/Time"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdAddTitle 
         Caption         =   "&Add"
         Height          =   375
         Left            =   3120
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtDescription 
         Height          =   555
         Left            =   1080
         TabIndex        =   17
         Top             =   720
         Width           =   4935
      End
      Begin VB.TextBox txtTitle 
         Height          =   435
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Transition"
         Height          =   195
         Left            =   1605
         TabIndex        =   42
         Top             =   5370
         Width           =   690
      End
      Begin VB.Label lblAutodisplay 
         AutoSize        =   -1  'True
         Caption         =   "Autodisplay is now disabled"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2520
         TabIndex        =   31
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label lblptrTitle 
         AutoSize        =   -1  'True
         Caption         =   "Pointer-0"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   5280
         TabIndex        =   30
         Top             =   4920
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Description"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Title"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileAdd 
         Caption         =   "A&dd New"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDispnow 
         Caption         =   "Display Selected N&ow"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "Dele&te"
      End
   End
End
Attribute VB_Name = "frmConfigurations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hi all, i'm so happy to present this code i was experimenting on
'how to link DHTML and vb6 and this is my output
'if u r on search of some code to manage your tv shows
'i hope this will help you
'i'm eagerly waiting for your comments so plz do write some feedback
Dim ptrTitle As Integer
Dim ed(17) As String

Private Sub chkEnableBN_Click()
'Display the breaking news
frmOutput.wbotp.Document.Form1.board5.Value = chkEnableBN.Value
End Sub

Private Sub chkTitle_Click()
'to enable/disable the autoscrolling of the title news
Timer1.Enabled = Not Timer1.Enabled
lblAutodisplay.Caption = "Auto display is now " & IIf(Timer1.Enabled, "Enabled", "Disabled")
End Sub


Private Sub cmbEffects_Click()
'apply the selected transition
DoEvents
frmOutput.wbotp.Document.Form1.board4.Value = cmbEffects.List(cmbEffects.ListIndex) & "(" & ed(cmbEffects.ListIndex)
End Sub

Private Sub cmbEffects_GotFocus()
'load the transitions from a file
'make sure that the file exists
On Error GoTo er
If cmbEffects.ListCount = 0 Then
    Dim infile As Integer
    Dim str As String, data As String
    infile = FreeFile
    Open App.Path & "\effects.txt" For Input As infile
        While Not EOF(infile)
            Line Input #infile, str
            data = data & str & vbCrLf
        Wend
        Close #infile
        cmbEffects.Clear
        ab = Split(data, vbCrLf)
        For i = 0 To UBound(ab) - 1
            abc = Split(ab(i), "(")
            cmbEffects.AddItem abc(0)
            ed(i) = abc(1)
        Next
End If
er:
End Sub

Private Sub cmdAddTitle_Click()
'to add up new titles
Dim li As ListItem
If Trim(txtTitle) = "" Then
    MsgBox "Enter a valid title", vbExclamation, "Error"
    txtTitle.SetFocus
ElseIf Trim(txtDescription) = "" Then
    MsgBox "Enter a valid description", vbExclamation, "Error"
    txtDescription.SetFocus
Else
    Set li = lstTitles.ListItems.Add(, , lstTitles.ListItems.Count + 1)
        li.ListSubItems.Add , , Trim(txtTitle.Text)
        li.ListSubItems.Add , , Trim(txtDescription.Text)
        li.ListSubItems.Add , , Now()
    txtTitle.Text = ""
    txtDescription.Text = ""
    txtTitle.SetFocus
End If
End Sub


Private Sub cmdApply_Click()
'nothing
MsgBox "Press relevent buttons to apply configurations"
End Sub

Private Sub cmdCancel_Click()
'hey plz post some feedback yaar..!!
frmFeedback.Show vbModal, Me
End Sub

Private Sub cmdClearTitle_Click()
txtTitle.Text = ""
txtDescription.Text = ""
End Sub

Private Sub cmdDeleteTitle_Click()
'of course to delete the titles
Dim i As Integer, c As Integer
Dim li As ListItem
For i = 1 To lstTitles.ListItems.Count
    If lstTitles.ListItems(i).Checked Then c = c + 1
Next
If lstTitles.ListItems.Count = 0 Or c <= 0 Then Exit Sub
If MsgBox("Are you sure to delete " & c & " itmes?", vbQuestion + vbYesNo, "Confirm Delete.") = vbNo Then Exit Sub
i = 0
While i < lstTitles.ListItems.Count
    i = i + 1
    Set li = lstTitles.ListItems(i)
    If li.Checked Then
        lstTitles.ListItems.Remove (i)
        i = 0
    End If
Wend
For i = 1 To lstTitles.ListItems.Count
    lstTitles.ListItems(i).Text = i
Next
End Sub


Private Sub cmdOK_Click()
'initialize
Dim ctr As Control
For Each ctr In Me.Controls
    Debug.Print ctr.Name
    ctr.Enabled = Not InStr(1, ctr.Name, "mnusep")
 Next ctr
cmdOK.Enabled = Not cmdOK.Enabled
Timer1.Enabled = False
lstTitles_Click

'this loads the output screen
frmOutput.Show
'load the html page where the DHTML code resides
frmOutput.wbotp.Navigate App.Path & "\effects.htm"
'wait for the page to fully load  then do the manipulations
Timer2.Enabled = True
End Sub

Private Sub cmdScrollAdd_Click()
'entry of scrolling news
Dim li As ListItem
If Trim(txtScrolltitle.Text) = "" Then
    MsgBox "Enter a valid title", vbExclamation, "Error"
    txtScrolltitle.SetFocus
ElseIf Trim(txtScrollDesc) = "" Then
    MsgBox "Enter a valid description", vbExclamation, "Error"
    txtScrollDesc.SetFocus
Else
    Set li = lstScroll.ListItems.Add(, , lstScroll.ListItems.Count + 1)
        li.ListSubItems.Add , , Trim(txtScrolltitle.Text)
        li.ListSubItems.Add , , Trim(txtScrollDesc.Text)
        li.ListSubItems.Add , , Now()
    txtScrolltitle.Text = ""
    txtScrollDesc.Text = ""
    txtScrolltitle.SetFocus
End If
End Sub

Private Sub cmdScrollClear_Click()
txtScrolltitle.Text = ""
txtScrollDesc.Text = ""
txtScrolltitle.SetFocus
End Sub

Private Sub cmdScrollDelete_Click()
Dim i As Integer, c As Integer
Dim li As ListItem
For i = 1 To lstScroll.ListItems.Count
    If lstScroll.ListItems(i).Checked Then c = c + 1
Next
If lstScroll.ListItems.Count = 0 Or c <= 0 Then Exit Sub
If MsgBox("Are you sure to delete " & c & " itmes?", vbQuestion + vbYesNo, "Confirm Delete.") = vbNo Then Exit Sub
i = 0
While i < lstScroll.ListItems.Count
    i = i + 1
    Set li = lstScroll.ListItems(i)
    If li.Checked Then
        lstScroll.ListItems.Remove (i)
        i = 0
    End If
Wend
For i = 1 To lstScroll.ListItems.Count
    lstScroll.ListItems(i).Text = i
Next
End Sub

Private Sub cmdScrollSet_Click()
'set the new scrolling news
Dim li As ListItem, scr As String
For i = 1 To lstScroll.ListItems.Count
    Set li = lstScroll.ListItems(i)
    scr = scr & li.ListSubItems(2).Text & " " & "<img src=" & imgsep.Tag & "></img>"
Next
    frmOutput.wbotp.Document.Form1.board3.Value = scr
End Sub



Private Sub cmdsetBrknews_Click()
frmOutput.wbotp.Document.Form1.board2.Value = Trim(txtBrknews.Text)
lblBN.Caption = "Breaking news is set to :" & vbCrLf & Trim(txtBrknews.Text)
End Sub

Private Sub cmdShootTitle_Click()
    frmOutput.wbotp.Document.Form1.board1.Value = Trim(txtTitle.Text) & "<v>" & Trim(txtDescription.Text)
End Sub


Private Sub Form_Load()
Dim ctr As Control
'load the saved titles
If Dir(App.Path & "\temps.txt") <> "" Then
    read_file App.Path & "\temps.txt", lstTitles
End If
'load the saved scrolling news
If Dir(App.Path & "\tempscroll.txt") <> "" Then
    read_file App.Path & "\tempscroll.txt", lstScroll
End If
'load the saved breaking news
txtBrknews = GetSetting("NM_pro", "BNews", "Title")

'disable the controls until the output form gets loaded
For Each ctr In Me.Controls
    Debug.Print ctr.Name
    ctr.Enabled = InStr(1, ctr.Name, "mnusep")
Next ctr
'except the load output button
cmdOK.Enabled = Not cmdOK.Enabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
'save all
save_values App.Path & "\temps.txt", lstTitles
save_values App.Path & "\tempscroll.txt", lstScroll
SaveSetting "NM_pro", "BNews", "Title", Trim(txtBrknews)
End
End Sub


Private Sub lstScroll_Click()
If lstScroll.ListItems.Count = 0 Then Exit Sub
Dim li As ListItem
Set li = lstScroll.SelectedItem
    txtScrolltitle.Text = li.ListSubItems(1).Text
    txtScrollDesc.Text = li.ListSubItems(2).Text
End Sub

Private Sub lstTitles_Click()
If lstTitles.ListItems.Count = 0 Then Exit Sub
Dim li As ListItem
Set li = lstTitles.SelectedItem
    txtTitle.Text = li.ListSubItems(1).Text
    txtDescription.Text = li.ListSubItems(2).Text
End Sub

Private Sub mnuFileAdd_Click()
txtTitle.Text = ""
txtDescription.Text = ""
txtTitle.SetFocus
End Sub

Private Sub mnuFileDelete_Click()
cmdDeleteTitle_Click
End Sub

Private Sub mnuFileDispnow_Click()
ptrTitle = lstTitles.SelectedItem.Index - 1
End Sub

Private Sub optBN_Click()
frameScrollingNews.Visible = False
frameBrknews.Visible = True
frameTitleNews.Visible = False
End Sub

Private Sub optScroll_Click()
frameScrollingNews.Visible = True
frameBrknews.Visible = False
frameTitleNews.Visible = False
End Sub

Private Sub optTitle_Click()
frameScrollingNews.Visible = False
frameBrknews.Visible = False
frameTitleNews.Visible = True
End Sub

Private Sub Timer1_Timer()
'to auto display title news
If lstTitles.ListItems.Count = 0 Then Exit Sub
    Dim li As ListItem
    ptrTitle = ptrTitle + 1
    If ptrTitle > lstTitles.ListItems.Count Then ptrTitle = 1
    Set li = lstTitles.ListItems(ptrTitle)
       DoEvents
       frmOutput.wbotp.Document.Form1.board1.Value = li.ListSubItems(1).Text & "<v>" & li.ListSubItems(2).Text
       DoEvents
       lblptrTitle.Caption = "Pointer--> " & ptrTitle
End Sub
Private Function save_values(filename As String, lst As ListView)
'to save list
    Dim li As ListItem, lis As ListSubItem
    Dim str As String
    
    For Each li In lst.ListItems
        str = str & "<tr>" & li.Text
       For Each lis In li.ListSubItems
            str = str & "<td>" & lis.Text
       Next
    Next
    Open filename For Output As #1
        Print #1, str
    Close #1
End Function
Private Function read_file(filename As String, lst As ListView)
Dim li As ListItem, infile As Integer
Dim str As String
'to reload saved lists
infile = FreeFile
Open filename For Input As infile
    While Not EOF(infile)
        Line Input #infile, str
    Wend
    Close #infile
    ab = Split(str, "<tr>")
    For i = 1 To UBound(ab)
        abc = Split(ab(i), "<td>")
        Set li = lst.ListItems.Add(, , abc(0))
        For j = 1 To UBound(abc)
            li.ListSubItems.Add , , abc(j)
        Next
    Next
End Function

Private Sub Timer2_Timer()
If frmOutput.wbotp.Busy = False Then
    cmdShootTitle_Click
    cmdsetBrknews_Click
    cmdScrollSet_Click
    Timer2.Enabled = False
End If
End Sub
