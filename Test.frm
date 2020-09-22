VERSION 5.00
Begin VB.Form fTest 
   Caption         =   "Light toolbar 2.1 test"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkToolbarsEnabled 
      Caption         =   "Toolbars enabled"
      Height          =   270
      Left            =   4110
      TabIndex        =   0
      Top             =   4455
      Value           =   1  'Checked
      Width           =   1560
   End
   Begin prjTest.ucToolbar ucToolbar3 
      Height          =   495
      Left            =   1320
      Top             =   1320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      BackColor       =   8438015
      BarEdge         =   -1  'True
   End
   Begin prjTest.ucToolbar ucToolbar1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   714
      BarEdge         =   -1  'True
   End
   Begin prjTest.ucToolbar ucToolbar2 
      Height          =   3090
      Left            =   120
      Top             =   600
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   5450
      BarOrientation  =   1
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()

    '-- Button type = [NORMAL]
    ucToolbar1.BuildToolbar LoadResPicture("TB_MAIN", vbResBitmap), &HFF00FF, 16, "NNN|NN|NNN|N|N"
    ucToolbar1.SetTooltips "New|Load|Save|Undo|Redo|Cut|Copy|Paste|Screen capture|Set hot spot"
    
    '-- Button type = [OPTION]
    ucToolbar2.BuildToolbar LoadResPicture("TB_TOOLS", vbResBitmap), &HFF00FF, 20, "OOOOOOOOOO"
    ucToolbar2.SetTooltips "Selection frame|Pencil|Straight line|Brush|Flood fill|Color eraser|Shape|Text|Color selector|Color locator"
    
    '-- Button type = [NORMAL+CHECK+OPTION] / No tooltips
    ucToolbar3.BuildToolbar LoadResPicture("TB_TEXT", vbResBitmap), &HFF00FF, 16, "CCC|OOO|NN|NN"
    
    '-- Test: selecting option buttons
    ucToolbar2.CheckButton 1, True
    ucToolbar3.CheckButton 4, True
End Sub

Private Sub chkToolbarsEnabled_Click()
    ucToolbar1.Enabled = -chkToolbarsEnabled
    ucToolbar2.Enabled = -chkToolbarsEnabled
    ucToolbar3.Enabled = -chkToolbarsEnabled
End Sub



Private Sub ucToolbar1_ButtonClick(ByVal Index As Long, ByVal MouseButton As Integer, ByVal xLeft As Long, ByVal yTop As Long)
    Debug.Print "ucToolbar1_ButtonClick"; Index; MouseButton
End Sub
Private Sub ucToolbar1_ButtonCheck(ByVal Index As Long, ByVal xLeft As Long, ByVal yTop As Long)
    Debug.Print "ucToolbar1_ButtonCheck"; Index
End Sub

Private Sub ucToolbar2_ButtonClick(ByVal Index As Long, ByVal MouseButton As Integer, ByVal xLeft As Long, ByVal yTop As Long)
    Debug.Print "ucToolbar2_ButtonClick"; Index; MouseButton
End Sub
Private Sub ucToolbar2_ButtonCheck(ByVal Index As Long, ByVal xLeft As Long, ByVal yTop As Long)
    Debug.Print "ucToolbar2_ButtonCheck"; Index
End Sub

Private Sub ucToolbar3_ButtonCheck(ByVal Index As Long, ByVal xLeft As Long, ByVal yTop As Long)
    Debug.Print "ucToolbar3_ButtonCheck"; Index
End Sub
Private Sub ucToolbar3_ButtonClick(ByVal Index As Long, ByVal MouseButton As Integer, ByVal xLeft As Long, ByVal yTop As Long)
    Debug.Print "ucToolbar3_ButtonClick"; Index; MouseButton
End Sub
