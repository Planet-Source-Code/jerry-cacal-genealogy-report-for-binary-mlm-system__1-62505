VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Binary MLM System"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   14535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   " MEMBERS' INFORMATION "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   14295
      Begin VB.CommandButton Command1 
         Caption         =   "Generate Genealogy Report"
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   3360
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Count Downlines"
         Height          =   495
         Left            =   3120
         TabIndex        =   2
         Top             =   3360
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Select a member from the list and click on the buttons below to view the genealogy or count the downlines"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   10575
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Close"
      Height          =   495
      Left            =   12360
      TabIndex        =   6
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   7680
      Width           =   6135
      Begin VB.CommandButton Command3 
         Caption         =   "Find Members with no Upline"
         Height          =   495
         Left            =   3480
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "To view members with no connection in the binary structure or no upline, click here..."
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   6480
      TabIndex        =   9
      Top             =   7680
      Width           =   5655
      Begin VB.CommandButton Command4 
         Caption         =   "Refresh Data"
         Height          =   495
         Left            =   3600
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "If you modified the database, click this button to refresh the data in the above listview"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " GENEALOGY REPORT "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   14295
      Begin MSComctlLib.ListView ListView2 
         Height          =   2295
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Caption         =   "GENEALOGY FOR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   14055
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DefList()

  ListView1.ListItems.Clear
  ListView1.ColumnHeaders.Add 1, "K1", "rKey"             ' ID Key
  ListView1.ColumnHeaders.Add 2, "K2", "Date Joined"      ' The date when the member joined
  ListView1.ColumnHeaders.Add 3, "K3", "MemberCode"       ' Member ID Code
  ListView1.ColumnHeaders.Add 4, "K4", "Name"             ' Name
  ListView1.ColumnHeaders.Add 5, "K5", "Address"          ' Address
  ListView1.ColumnHeaders.Add 6, "K6", "Cellphone"        ' Mobile phone no.
  ListView1.ColumnHeaders.Add 7, "K7", "Direct Referral"  ' The ID Code of member who referred this member - must pre-exist when this member is entered
  ListView1.ColumnHeaders.Add 8, "K8", "Upline"           ' The ID Code of member who is the upline of this member - must pre-exist when this member is entered
  ListView1.ColumnHeaders.Add 9, "K9", "Binary Position"  ' Left or Right connection with the upline
  ListView1.ColumnHeaders.Add 10, "K10", "Status"         ' Active or Inactive status
  
  ListView2.ListItems.Clear
  ListView2.ColumnHeaders.Add 1, "K1", "Wing"             ' Member's position in reference with the center or the topmost member
  ListView2.ColumnHeaders.Add 2, "K2", "Date Joined"      ' Date Joined
  ListView2.ColumnHeaders.Add 3, "K3", "MemberCode"       ' Member ID Code
  ListView2.ColumnHeaders.Add 4, "K4", "Name"             ' Name
  ListView2.ColumnHeaders.Add 5, "K5", "DR Code"          ' Direct Referral Code
  ListView2.ColumnHeaders.Add 6, "K6", "DR Name"          ' Direct Referral Name
  ListView2.ColumnHeaders.Add 7, "K7", "Upline Code"      ' Upline Code
  ListView2.ColumnHeaders.Add 8, "K8", "Upline Name"      ' Upline Name
  ListView2.ColumnHeaders.Add 9, "K9", "Binary Position"  ' Left or Right connection with upline
  ListView2.ColumnHeaders.Add 10, "K10", "Status"         ' Active or Inactive status

End Sub

Private Sub FillList()

  Dim txtLine As Variant
  Dim X As Long
  Dim Y As Integer
  
  txtLine = FileToArray(sTempFileDir)
  ListView1.ListItems.Clear
  ListView1.Visible = False
  For X = 0 To UBound(txtLine) - 1 Step 1
    With ListView1
      .ListItems.Add , "K" & X + 1, ParamValue("|", txtLine(X), 1)
      For Y = 2 To ParamCount("|", txtLine(X)) Step 1
        .ListItems("K" & X + 1).SubItems(Y - 1) = ParamValue("|", txtLine(X), Y)
      Next Y
    End With
  Next X
  ListView1.Visible = True

End Sub

Private Sub Command1_Click()

  Dim GeneaReport As Variant
  Dim MemberCode As String
  Dim TotalLeftWing As Long, TotalRightWing As Long, TotalLevels As Long
  Dim X As Long, Y As Integer
  
  MemberCode = ListView1.SelectedItem.ListSubItems(2)
  If MemberCodeExist(MemberCode) Then
    frmMain.Enabled = False
    frmWait.Show 0, frmMain
    
    GeneaReport = GetGenealogy(MemberCode, TotalLeftWing, TotalRightWing, TotalLevels)
    
    ListView2.Visible = False
    ListView2.ListItems.Clear
    
    For X = 0 To UBound(GeneaReport) - 1 Step 1
      With ListView2
        .ListItems.Add , "K" & X + 1, ParamValue("|", GeneaReport(X), 1)
        For Y = 2 To ParamCount("|", GeneaReport(X)) Step 1
          .ListItems("K" & X + 1).SubItems(Y - 1) = ParamValue("|", GeneaReport(X), Y)
        Next Y
      End With
    Next X
    
    ListView2.Visible = True
    Unload frmWait
    frmMain.Enabled = True
    Label1.Caption = "Genealogy for : " & MemberCode & " " & ListView1.SelectedItem.ListSubItems(3)
    MsgBox "Total downlines for " & MemberCode & " " & ListView1.SelectedItem.ListSubItems(3) & vbCrLf & vbCrLf & "Total Left Wing : " & TotalLeftWing & vbCrLf & "Total Right Wing : " & TotalRightWing & vbCrLf & "Total Levels : " & TotalLevels, vbOKOnly + vbInformation, "Total Downlines"
  Else
    MsgBox "Member Code " & MemberCode & " not found."
  End If

End Sub

Private Sub Command2_Click()

  Dim CountMemberCode As String
  
  CountMemberCode = ListView1.SelectedItem.ListSubItems(2)
  
  If MemberCodeExist(CountMemberCode) Then
    frmMain.Enabled = False
    frmWait.Show 0, frmMain
    CountDownlines CountMemberCode
    Unload frmWait
    frmMain.Enabled = True
    MsgBox "Total downlines for " & CountMemberCode & " " & ListView1.SelectedItem.ListSubItems(3) & vbCrLf & vbCrLf & "Total Left Wing : " & TotalLeft & vbCrLf & "Total Right Wing : " & TotalRight, vbOKOnly + vbInformation, "Total Downlines"
    ' To know the grand total of downlines, just add the Total Left and Total Right.
  Else
    MsgBox "Member Code not found."
  End If

End Sub

Private Sub Command3_Click()

  FindNoUpline

End Sub

Private Sub Command4_Click()

  InitData
  BinaryInitialize
  FillList

End Sub

Private Sub Command5_Click()

  End

End Sub

Private Sub Form_Load()

  BinaryInitialize
  DefList
  FillList

End Sub
