VERSION 5.00
Begin VB.Form Issue_Form 
   Caption         =   "Rental Record"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   LinkTopic       =   "Form1"
   Picture         =   "Issue form.frx":0000
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame issue_report_frame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Report Issue Frame"
      Height          =   12375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   23055
      Begin VB.TextBox issue_fr_reporter_textbox 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         TabIndex        =   46
         Text            =   "Enter reporter's name"
         Top             =   5400
         Width           =   3855
      End
      Begin VB.TextBox issue_fr_issue_textbox 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   10440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Top             =   6840
         Width           =   3855
      End
      Begin VB.TextBox issue_fr_contact_textbox 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         TabIndex        =   44
         Text            =   "Enter contact number"
         Top             =   6120
         Width           =   3855
      End
      Begin VB.CommandButton issue_fr_report_command 
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8520
         TabIndex        =   2
         Top             =   9240
         Width           =   5775
      End
      Begin VB.ComboBox issue_fr_roomNum_combo 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "Issue form.frx":9D0A
         Left            =   10440
         List            =   "Issue form.frx":9D17
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "Room number"
         Top             =   4680
         Width           =   3855
      End
      Begin VB.Image Image22 
         Height          =   1050
         Left            =   5520
         Picture         =   "Issue form.frx":9D24
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label issue_fr_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Reporter's name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   8520
         TabIndex        =   9
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label issue_fr_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact number"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   8520
         TabIndex        =   8
         Top             =   6120
         Width           =   1395
      End
      Begin VB.Label issue_fr_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Issue"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   8520
         TabIndex        =   7
         Top             =   6840
         Width           =   435
      End
      Begin VB.Label issue_fr_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Room Number"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   8520
         TabIndex        =   6
         Top             =   4680
         Width           =   1305
      End
      Begin VB.Label issue_fr_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Report an issue"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   450
         Index           =   1
         Left            =   10260
         TabIndex        =   5
         Top             =   3360
         Width           =   2205
      End
      Begin VB.Label issue_fr_label 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Rental Record"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   36
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   720
         TabIndex        =   4
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label issue_fr_goback 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Go back"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   10920
         TabIndex        =   3
         Top             =   10200
         Width           =   885
      End
      Begin VB.Image Image14 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   8160
         Picture         =   "Issue form.frx":13A2E
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   6495
      End
      Begin VB.Image Image3 
         Height          =   615
         Left            =   8520
         Picture         =   "Issue form.frx":13F38
         Stretch         =   -1  'True
         Top             =   10080
         Width           =   5775
      End
      Begin VB.Image Image6 
         Height          =   9855
         Left            =   6240
         Picture         =   "Issue form.frx":14443
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   10335
      End
   End
   Begin VB.Frame issue_detail_frame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Issue Detail Frame"
      Height          =   12375
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   23055
      Begin VB.Frame issue_filter_frame 
         BorderStyle     =   0  'None
         Caption         =   "Tenant Filter Frame"
         Height          =   1575
         Left            =   15840
         TabIndex        =   36
         Top             =   10320
         Width           =   6375
         Begin VB.TextBox issue_serial_textbox 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3360
            TabIndex        =   42
            Text            =   "Enter serial number"
            Top             =   240
            Width           =   2775
         End
         Begin VB.CommandButton issue_solve_command 
            Caption         =   "Solve"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   3360
            TabIndex        =   41
            Top             =   960
            Width           =   2775
         End
         Begin VB.CheckBox issue_filter_main_option 
            Caption         =   "Issue Status"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   40
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton issue_filter_option 
            Caption         =   "Unsolved"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   1800
            TabIndex        =   38
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton issue_filter_filter_command 
            Caption         =   "Filter"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   240
            TabIndex        =   37
            Top             =   960
            Width           =   2775
         End
         Begin VB.OptionButton issue_filter_option 
            Caption         =   "Solved"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   1800
            TabIndex        =   39
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label issue_filter_filter_status 
            Caption         =   "Filter  : On"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   43
            Top             =   600
            Width           =   1215
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000000&
            Height          =   1575
            Left            =   0
            Top             =   0
            Width           =   3135
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H80000000&
            Height          =   1575
            Left            =   3120
            Top             =   0
            Width           =   3255
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   22200
         TabIndex        =   23
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   720
         TabIndex        =   22
         Top             =   3960
         Width           =   21495
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   720
         TabIndex        =   21
         Top             =   10080
         Width           =   21495
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   720
         TabIndex        =   20
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   720
         TabIndex        =   19
         Top             =   3480
         Width           =   21495
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   11040
         TabIndex        =   18
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   8760
         TabIndex        =   17
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   7080
         TabIndex        =   16
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   5520
         TabIndex        =   15
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   2640
         TabIndex        =   13
         Top             =   3480
         Width           =   15
         Begin VB.Frame Frame4 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   6735
            Left            =   120
            TabIndex        =   14
            Top             =   2520
            Width           =   135
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   1440
         TabIndex        =   12
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   13320
         TabIndex        =   11
         Top             =   3480
         Width           =   15
      End
      Begin VB.TextBox issue_detail_TextBox 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6015
         Left            =   840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   24
         Text            =   "Issue form.frx":1AB8F
         Top             =   4080
         Width           =   21375
      End
      Begin VB.Image Image25 
         Height          =   1050
         Left            =   5520
         Picture         =   "Issue form.frx":1AC22
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label issue_detail_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Reporter's Name"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2760
         TabIndex        =   35
         Top             =   3600
         Width           =   1800
      End
      Begin VB.Label issue_detail_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Room N0."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   1560
         TabIndex        =   34
         Top             =   3600
         Width           =   960
      End
      Begin VB.Label issue_detail_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Status"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   7200
         TabIndex        =   33
         Top             =   3600
         Width           =   1440
      End
      Begin VB.Label issue_detail_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Reported Date"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   8880
         TabIndex        =   32
         Top             =   3600
         Width           =   1560
      End
      Begin VB.Label issue_detail_goback 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Go Back"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1560
         TabIndex        =   31
         Top             =   11160
         Width           =   1035
      End
      Begin VB.Label issue_detail_label 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Rental Record"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   36
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   720
         TabIndex        =   30
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label issue_detail_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "S.N."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   840
         TabIndex        =   29
         Top             =   3600
         Width           =   480
      End
      Begin VB.Label issue_detail_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Tenant Details"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   18
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Index           =   1
         Left            =   960
         TabIndex        =   28
         Top             =   2520
         Width           =   2355
      End
      Begin VB.Label issue_detail_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact N0."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   5640
         TabIndex        =   27
         Top             =   3600
         Width           =   1320
      End
      Begin VB.Label issue_detail_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Solved Date"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   11160
         TabIndex        =   26
         Top             =   3600
         Width           =   1320
      End
      Begin VB.Label issue_detail_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Detail"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   13440
         TabIndex        =   25
         Top             =   3600
         Width           =   1440
      End
      Begin VB.Image Image17 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   720
         Picture         =   "Issue form.frx":2492C
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Image Image26 
         Height          =   495
         Left            =   720
         Picture         =   "Issue form.frx":24E36
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   21495
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   720
         Picture         =   "Issue form.frx":25341
         Stretch         =   -1  'True
         Top             =   11040
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Issue_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim today_issue As Date

Dim serial_issue As Integer

Dim counter_issue_1 As Integer
Dim counter_issue_2 As Integer

Dim new_issue As issue_class
Dim temp_issue As issue_class
Dim existing_issue As issue_class

Dim display_issue As display_issue_class

Private Sub issue_serial_textbox_GotFocus()
    If issue_serial_textbox.Text = "Enter serial number" Then issue_serial_textbox.Text = ""
End Sub

Private Sub issue_serial_textbox_LostFocus()
    If issue_serial_textbox.Text = "" Then issue_serial_textbox.Text = "Enter serial number"
End Sub

Private Sub issue_serial_textbox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        KeyAscii = KeyAscii
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        issue_solve_command_Click
    ElseIf KeyAscii < 47 Or KeyAscii > 58 Then
        KeyAscii = 0
    End If
End Sub


'solve command button
Private Sub issue_solve_command_Click()
    serial_issue = Val(issue_serial_textbox.Text)
    counter_issue_2 = issue_count_function
    temp_issue.issue_status = "Solved"
    'check if provided serial number is valid or not
    If serial_issue < 1 Or serial_issue > counter_issue_2 Then
        MsgBox_Response = MsgBox("                         Please enter the valid serial number.", vbInformation + vbOKOnly, "Rental Record")
        issue_serial_textbox.Text = ""
        issue_serial_textbox.SetFocus
    Else
        'solve
        Open "IssueDetail.txt" For Random As #1 Len = 167
            For counter_issue_1 = 1 To counter_issue_2 Step 1
                Get #1, counter_issue_1, existing_issue
                If serial_issue = counter_issue_1 Then
                    If existing_issue.issue_status = temp_issue.issue_status Then
                        MsgBox_Response = MsgBox("                    Sorry, this issue is already solved.", vbInformation + vbOKOnly, "Rental Record")
                    End If
                    existing_issue.issue_solved_date = Now
                    existing_issue.issue_status = "Solved"
                End If
                Put #1, counter_issue_1, existing_issue
            Next counter_issue_1
        Close #1
        
        'refresh textbox
        issue_detail_TextBox.Text = ""
        Open "IssueDetail.txt" For Random As #2 Len = 167
            For counter_issue_1 = 1 To counter_issue_2 Step 1
                Get #2, counter_issue_1, existing_issue
                display_issue.serial = CStr(counter_issue_1)
                display_issue.room_num = CStr(existing_issue.room_num)
                display_issue.reporter = existing_issue.reporter
                display_issue.contact_num = existing_issue.contact_num
                display_issue.issue_status = existing_issue.issue_status
                display_issue.issue_reported_date = Format(existing_issue.issue_reported_date, "dd mmmm, yyyy")
                
                If existing_issue.issue_solved_date = 0 Then
                    display_issue.issue_solved_date = "-"
                Else
                    display_issue.issue_solved_date = Format(existing_issue.issue_solved_date, "dd mmmm, yyyy")
                End If
                
                'display
                issue_detail_TextBox = issue_detail_TextBox + display_issue.serial + display_issue.room_num + display_issue.reporter + display_issue.contact_num
                issue_detail_TextBox = issue_detail_TextBox + display_issue.issue_status + display_issue.issue_reported_date + display_issue.issue_solved_date
                issue_detail_TextBox = issue_detail_TextBox + existing_issue.issue_detail + vbNewLine
            Next counter_issue_1
        Close #2
        
        'reset values
        issue_serial_textbox.Text = ""
        issue_filter_filter_status.Caption = "Filter : Off"
    End If
End Sub


Private Sub issue_detail_goback_Click()
    issue_serial_textbox.Text = "Enter serial number"
    Main_Form.WindowState = Issue_Form.WindowState
    Main_Form.Visible = True
    Issue_Form.Visible = False
End Sub

Private Sub issue_filter_filter_command_Click()
    If issue_filter_main_option.Value = Unchecked Then 'filter off
        issue_filter_filter_status.Caption = "Filter : Off"
        issue_detail_TextBox = ""
        counter_issue_2 = issue_count_function
        Open "IssueDetail.txt" For Random As #1 Len = 167
            For counter_issue_1 = 1 To counter_issue_2 Step 1
                Get #1, counter_issue_1, existing_issue
                display_issue.serial = CStr(counter_issue_1)
                display_issue.room_num = CStr(existing_issue.room_num)
                display_issue.reporter = existing_issue.reporter
                display_issue.contact_num = existing_issue.contact_num
                display_issue.issue_status = existing_issue.issue_status
                display_issue.issue_reported_date = Format(existing_issue.issue_reported_date, "dd mmmm, yyyy")
                
                If existing_issue.issue_solved_date = 0 Then
                    display_issue.issue_solved_date = "-"
                Else
                    display_issue.issue_solved_date = Format(existing_issue.issue_solved_date, "dd mmmm, yyyy")
                End If
                
                'display
                issue_detail_TextBox = issue_detail_TextBox + display_issue.serial + display_issue.room_num + display_issue.reporter + display_issue.contact_num + display_issue.issue_status
                issue_detail_TextBox = issue_detail_TextBox + display_issue.issue_reported_date + display_issue.issue_solved_date
                issue_detail_TextBox = issue_detail_TextBox + existing_issue.issue_detail + vbNewLine
            Next counter_issue_1
        Close #1
    ElseIf issue_filter_main_option.Value = Checked Then 'filter on
        temp_issue.issue_status = "Unsolved"
        
        issue_filter_filter_status.Caption = "Filter : On"
        issue_detail_TextBox = ""
        counter_issue_2 = issue_count_function
        
        Open "IssueDetail.txt" For Random As #1 Len = 167
            For counter_issue_1 = 1 To counter_issue_2 Step 1
                Get #1, counter_issue_1, existing_issue
                
                If issue_filter_option(1).Value = True Then 'unsolved issues
                     If existing_issue.issue_status = temp_issue.issue_status Then
                        display_issue.serial = CStr(counter_issue_1)
                        display_issue.room_num = CStr(existing_issue.room_num)
                        display_issue.reporter = existing_issue.reporter
                        display_issue.contact_num = existing_issue.contact_num
                        display_issue.issue_status = existing_issue.issue_status
                        display_issue.issue_reported_date = Format(existing_issue.issue_reported_date, "dd mmmm, yyyy")
                        display_issue.issue_solved_date = "-"
                        
                        'display
                        issue_detail_TextBox = issue_detail_TextBox + display_issue.serial + display_issue.room_num + display_issue.reporter + display_issue.contact_num + display_issue.issue_status
                        issue_detail_TextBox = issue_detail_TextBox + display_issue.issue_reported_date + display_issue.issue_solved_date
                        issue_detail_TextBox = issue_detail_TextBox + existing_issue.issue_detail + vbNewLine
                    End If
                Else
                    If existing_issue.issue_status <> temp_issue.issue_status Then
                        display_issue.serial = CStr(counter_issue_1)
                        display_issue.room_num = CStr(existing_issue.room_num)
                        display_issue.reporter = existing_issue.reporter
                        display_issue.contact_num = existing_issue.contact_num
                        display_issue.issue_status = existing_issue.issue_status
                        display_issue.issue_reported_date = Format(existing_issue.issue_reported_date, "dd mmmm, yyyy")
                        display_issue.issue_solved_date = Format(existing_issue.issue_solved_date, "dd mmmm, yyyy")
                                            
                        'display
                        issue_detail_TextBox = issue_detail_TextBox + display_issue.serial + display_issue.room_num + display_issue.reporter + display_issue.contact_num + display_issue.issue_status
                        issue_detail_TextBox = issue_detail_TextBox + display_issue.issue_reported_date + display_issue.issue_solved_date
                        issue_detail_TextBox = issue_detail_TextBox + existing_issue.issue_detail + vbNewLine
                    End If
                End If
            Next counter_issue_1
        Close #1
    End If
End Sub

Private Sub issue_fr_goback_Click()
    'reset textbox values
    issue_fr_roomNum_combo.Text = "Room number"
    issue_fr_reporter_textbox.Text = "Enter reporter's name"
    issue_fr_contact_textbox.Text = "Enter contact number"
    issue_fr_issue_textbox.Text = ""
    
    'form & frame visibility
    Main_Form.WindowState = Issue_Form.WindowState
    Main_Form.Visible = True
    Issue_Form.Visible = False
End Sub

Private Sub issue_fr_roomNum_combo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And issue_fr_roomNum_combo.Text <> "Room number" Then
        issue_fr_reporter_textbox.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub issue_fr_roomNum_combo_LostFocus()
    If issue_fr_roomNum_combo.Text = "" Then issue_fr_roomNum_combo.Text = "Room number"
End Sub


'reporter's name textbox
Private Sub issue_fr_reporter_textbox_GotFocus()
    If issue_fr_reporter_textbox.Text = "Enter reporter's name" Then issue_fr_reporter_textbox.Text = ""
End Sub

Private Sub issue_fr_reporter_textbox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If issue_fr_reporter_textbox.Text <> "Enter reporter's name" And issue_fr_reporter_textbox.Text <> "" Then
            KeyAscii = 0
            issue_fr_contact_textbox.SetFocus
        End If
    End If
End Sub

Private Sub issue_fr_reporter_textbox_LostFocus()
    If issue_fr_reporter_textbox.Text = "" Then issue_fr_reporter_textbox.Text = "Enter reporter's name"
End Sub


'contact number textbox
Private Sub issue_fr_contact_textbox_GotFocus()
    If issue_fr_contact_textbox.Text = "Enter contact number" Then issue_fr_contact_textbox.Text = ""
End Sub

Private Sub issue_fr_contact_textbox_LostFocus()
    If issue_fr_contact_textbox.Text = "" Then issue_fr_contact_textbox.Text = "Enter contact number"
End Sub

Private Sub issue_fr_contact_textbox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If issue_fr_contact_textbox.Text <> "" Then
            KeyAscii = 0
            issue_fr_issue_textbox.SetFocus
        End If
    ElseIf KeyAscii = 8 Then
        KeyAscii = KeyAscii
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    Else
        KeyAscii = KeyAscii
    End If
End Sub


'report command
Private Sub issue_fr_report_command_Click()
    'check for default values
    If issue_fr_roomNum_combo.Text = "Room number" Or issue_fr_reporter_textbox.Text = "Enter reporter's name" Or issue_fr_contact_textbox.Text = "Enter contact number" Or issue_fr_issue_textbox.Text = "" Then
        If issue_fr_roomNum_combo.Text = "Room number" Then
            MsgBox_Response = MsgBox("                   Please choose the room number first.", vbInformation + vbOKOnly, "Rental Record")
            issue_fr_roomNum_combo.SetFocus
        ElseIf issue_fr_reporter_textbox.Text = "Enter reporter's name" Then
            MsgBox_Response = MsgBox("                   Please enter the reporter's name first.", vbInformation + vbOKOnly, "Rental Record")
            issue_fr_reporter_textbox.SetFocus
        ElseIf issue_fr_contact_textbox.Text = "Enter contact number" Then
            MsgBox_Response = MsgBox("                   Please enter the contact number first.", vbInformation + vbOKOnly, "Rental Record")
            issue_fr_contact_textbox.SetFocus
        ElseIf issue_fr_issue_textbox.Text = "" Then
            MsgBox_Response = MsgBox("                   Please enter about an issue in detail first.", vbInformation + vbOKOnly, "Rental Record")
            issue_fr_issue_textbox.SetFocus
        End If
        Exit Sub
    End If
    
    issue_today = Now
    new_issue.room_num = Val(issue_fr_roomNum_combo.Text)
    new_issue.reporter = issue_fr_reporter_textbox.Text
    new_issue.contact_num = issue_fr_contact_textbox.Text
    new_issue.issue_status = "Unsolved"
    new_issue.issue_reported_date = issue_today
    new_issue.issue_solved_date = 0
    new_issue.issue_detail = issue_fr_issue_textbox.Text
    
    'write issue in a file
    counter_issue_1 = issue_count_function
    counter_issue_1 = counter_issue_1 + 1
    Open "IssueDetail.txt" For Random As #1 Len = 167
        Put #1, counter_issue_1, new_issue
    Close #1
    
    'reset textbox values
    issue_fr_roomNum_combo.Text = "Room number"
    issue_fr_reporter_textbox.Text = "Enter reporter's name"
    issue_fr_contact_textbox.Text = "Enter contact number"
    issue_fr_issue_textbox.Text = ""
    
    MsgBox_Response = MsgBox("     Issue recorded. Do you want to report another issue?", vbInformation + vbYesNo, "Rental Record")
    If MsgBox_Response = 6 Then 'yes
        issue_fr_roomNum_combo.SetFocus
    Else
        issue_fr_goback_Click
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unload Main_Form
    Unload Tenant_Form
    Unload Payment_Form
End Sub
