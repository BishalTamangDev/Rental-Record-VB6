VERSION 5.00
Begin VB.Form Main_Form 
   BackColor       =   &H80000005&
   Caption         =   "Rental Record"
   ClientHeight    =   12135
   ClientLeft      =   1950
   ClientTop       =   -2355
   ClientWidth     =   22800
   FillStyle       =   2  'Horizontal Line
   LinkTopic       =   "Form1"
   ScaleHeight     =   606.75
   ScaleMode       =   2  'Point
   ScaleWidth      =   1140
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame main_menu_frame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "main_menu_frame"
      Height          =   12495
      Left            =   0
      TabIndex        =   109
      Top             =   0
      Width           =   23055
      Begin VB.Frame payment_option_frame 
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         Height          =   4335
         Left            =   13800
         TabIndex        =   147
         Top             =   6480
         Width           =   4815
         Begin VB.Label payment_option_new_electricity_fee 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New Electricity Charge Rate"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   165
            Top             =   3600
            Width           =   3570
         End
         Begin VB.Image payment_option_image 
            Height          =   855
            Index           =   4
            Left            =   0
            Picture         =   "Rent Record.frx":0000
            Stretch         =   -1  'True
            Top             =   3360
            Width           =   4815
         End
         Begin VB.Label payment_option_new_service 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New Service Charge"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            TabIndex        =   164
            Top             =   2760
            Width           =   2565
         End
         Begin VB.Label payment_option_service 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Charge"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   163
            Top             =   1920
            Width           =   1905
         End
         Begin VB.Label payment_option_report 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Show Reports"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   149
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label payment_option_payment 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Payment"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1785
            TabIndex        =   148
            Top             =   1080
            Width           =   1140
         End
         Begin VB.Image payment_option_image 
            Height          =   855
            Index           =   0
            Left            =   0
            Picture         =   "Rent Record.frx":050B
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4815
         End
         Begin VB.Image payment_option_image 
            Height          =   855
            Index           =   1
            Left            =   0
            Picture         =   "Rent Record.frx":0A16
            Stretch         =   -1  'True
            Top             =   840
            Width           =   4815
         End
         Begin VB.Image payment_option_image 
            Height          =   855
            Index           =   2
            Left            =   0
            Picture         =   "Rent Record.frx":0F21
            Stretch         =   -1  'True
            Top             =   1680
            Width           =   4815
         End
         Begin VB.Image payment_option_image 
            Height          =   855
            Index           =   3
            Left            =   0
            Picture         =   "Rent Record.frx":142C
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   4815
         End
         Begin VB.Shape Shape7 
            BorderColor     =   &H00C0C0C0&
            Height          =   4215
            Left            =   0
            Top             =   0
            Width           =   4815
         End
      End
      Begin VB.Frame tenant_option_frame 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   13800
         TabIndex        =   142
         Top             =   5640
         Width           =   4815
         Begin VB.Shape Shape5 
            BorderColor     =   &H00C0C0C0&
            Height          =   3375
            Left            =   0
            Top             =   0
            Width           =   4815
         End
         Begin VB.Label tenant_option_frame_add 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Add New Tenant"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1305
            TabIndex        =   146
            Top             =   1080
            Width           =   2145
         End
         Begin VB.Label tenant_option_frame_remove 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remove Tenant"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            TabIndex        =   145
            Top             =   1920
            Width           =   2010
         End
         Begin VB.Label tenant_option_frame_edit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Edit Tenant Details"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            TabIndex        =   144
            Top             =   2760
            Width           =   2415
         End
         Begin VB.Label tenant_option_frame_view 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "View Tenant Details"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            TabIndex        =   143
            Top             =   240
            Width           =   2535
         End
         Begin VB.Image tenant_option_image 
            Height          =   855
            Index           =   0
            Left            =   0
            Picture         =   "Rent Record.frx":1937
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4815
         End
         Begin VB.Image tenant_option_image 
            Height          =   855
            Index           =   1
            Left            =   0
            Picture         =   "Rent Record.frx":1E42
            Stretch         =   -1  'True
            Top             =   840
            Width           =   4815
         End
         Begin VB.Image tenant_option_image 
            Height          =   855
            Index           =   2
            Left            =   0
            Picture         =   "Rent Record.frx":234D
            Stretch         =   -1  'True
            Top             =   1680
            Width           =   4815
         End
         Begin VB.Image tenant_option_image 
            Height          =   855
            Index           =   3
            Left            =   0
            Picture         =   "Rent Record.frx":2858
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   4815
         End
      End
      Begin VB.Frame room_option_frame 
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   13800
         TabIndex        =   137
         Top             =   4800
         Width           =   4815
         Begin VB.Shape Shape6 
            BorderColor     =   &H00C0C0C0&
            Height          =   3375
            Left            =   0
            Top             =   0
            Width           =   4815
         End
         Begin VB.Label room_option_frame_view 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "View Room Details"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            TabIndex        =   141
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label room_option_frame_edit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Edit Room Details"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            TabIndex        =   140
            Top             =   2760
            Width           =   2295
         End
         Begin VB.Image room_option_image 
            Height          =   855
            Index           =   3
            Left            =   0
            Picture         =   "Rent Record.frx":2D63
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   4815
         End
         Begin VB.Label room_option_frame_remove 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remove Room"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   139
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label room_option_frame_add 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Add New Room"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   138
            Top             =   1080
            Width           =   2055
         End
         Begin VB.Image room_option_image 
            Height          =   855
            Index           =   0
            Left            =   0
            Picture         =   "Rent Record.frx":326E
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4815
         End
         Begin VB.Image room_option_image 
            Height          =   855
            Index           =   1
            Left            =   0
            Picture         =   "Rent Record.frx":3779
            Stretch         =   -1  'True
            Top             =   840
            Width           =   4815
         End
         Begin VB.Image room_option_image 
            Height          =   855
            Index           =   2
            Left            =   0
            Picture         =   "Rent Record.frx":3C84
            Stretch         =   -1  'True
            Top             =   1680
            Width           =   4815
         End
      End
      Begin VB.Frame issue_option_frame 
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   13800
         TabIndex        =   150
         Top             =   7320
         Width           =   4815
         Begin VB.Shape Shape2 
            BorderColor     =   &H00C0C0C0&
            Height          =   1695
            Left            =   0
            Top             =   0
            Width           =   4815
         End
         Begin VB.Label issue_option_report 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Report Issue"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   152
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label issue_option_details 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Details"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   151
            Top             =   240
            Width           =   975
         End
         Begin VB.Image issue_option_image 
            Height          =   855
            Index           =   1
            Left            =   0
            Picture         =   "Rent Record.frx":418F
            Stretch         =   -1  'True
            Top             =   840
            Width           =   4815
         End
         Begin VB.Image issue_option_image 
            Height          =   855
            Index           =   0
            Left            =   0
            Picture         =   "Rent Record.frx":469A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4815
         End
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00E0E0E0&
         Height          =   4215
         Left            =   9000
         Top             =   4800
         Width           =   4815
      End
      Begin VB.Image Image18 
         Height          =   1050
         Left            =   5520
         Picture         =   "Rent Record.frx":4BA5
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label main_manu_note_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note : Welcome To Rental Record"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   840
         TabIndex        =   127
         Top             =   11040
         Width           =   3480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         Height          =   195
         Left            =   1200
         TabIndex        =   126
         Top             =   4320
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         Height          =   195
         Left            =   720
         TabIndex        =   125
         Top             =   4320
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label main_menu_logout_option 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Log Out"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10875
         TabIndex        =   124
         Top             =   8400
         Width           =   1020
      End
      Begin VB.Label main_menu_issue_option 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11055
         TabIndex        =   123
         Top             =   7560
         Width           =   645
      End
      Begin VB.Label main_menu_payment_option 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10830
         TabIndex        =   122
         Top             =   6720
         Width           =   1140
      End
      Begin VB.Label main_menu_tenant_option 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tenant"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10935
         TabIndex        =   121
         Top             =   5880
         Width           =   885
      End
      Begin VB.Label main_menu_room_option 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11010
         TabIndex        =   120
         Top             =   5040
         Width           =   765
      End
      Begin VB.Image option_image 
         Appearance      =   0  'Flat
         Height          =   855
         Index           =   0
         Left            =   8993
         Picture         =   "Rent Record.frx":E8AF
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   4815
      End
      Begin VB.Label main_menu_label 
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
         TabIndex        =   110
         Top             =   480
         Width           =   4695
      End
      Begin VB.Image option_image 
         Height          =   855
         Index           =   1
         Left            =   9000
         Picture         =   "Rent Record.frx":EDBA
         Stretch         =   -1  'True
         Top             =   5640
         Width           =   4815
      End
      Begin VB.Image option_image 
         Height          =   855
         Index           =   4
         Left            =   9000
         Picture         =   "Rent Record.frx":F2C5
         Stretch         =   -1  'True
         Top             =   8160
         Width           =   4815
      End
      Begin VB.Image option_image 
         Height          =   855
         Index           =   3
         Left            =   9000
         Picture         =   "Rent Record.frx":F7D0
         Stretch         =   -1  'True
         Top             =   7320
         Width           =   4815
      End
      Begin VB.Image option_image 
         Height          =   855
         Index           =   2
         Left            =   9000
         Picture         =   "Rent Record.frx":FCDB
         Stretch         =   -1  'True
         Top             =   6480
         Width           =   4815
      End
      Begin VB.Label main_menu_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Main Menu"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   1
         Left            =   960
         TabIndex        =   111
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Image Image10 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   720
         Picture         =   "Rent Record.frx":101E6
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   21615
      End
      Begin VB.Image Image9 
         Height          =   570
         Left            =   720
         Picture         =   "Rent Record.frx":106F0
         Stretch         =   -1  'True
         Top             =   10920
         Width           =   21585
      End
   End
   Begin VB.Frame add_room_frame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Add Room Frame"
      Height          =   12375
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Width           =   23055
      Begin VB.CheckBox add_room_check1 
         BackColor       =   &H8000000E&
         Caption         =   "Electricity"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   10560
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   78
         Top             =   6720
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox add_room_check1 
         BackColor       =   &H8000000E&
         Caption         =   "Security"
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   10560
         TabIndex        =   77
         Top             =   8280
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox add_room_check1 
         BackColor       =   &H8000000E&
         Caption         =   "Waste Management"
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   10560
         TabIndex        =   76
         Top             =   7800
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox add_room_check1 
         BackColor       =   &H8000000E&
         Caption         =   "Drinking Water"
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   10560
         TabIndex        =   75
         Top             =   7320
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.TextBox add_room_rent_textbox 
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
         Left            =   10560
         TabIndex        =   74
         Text            =   "Enter rent amount"
         Top             =   5280
         Width           =   3855
      End
      Begin VB.TextBox add_room_bhk_textbox 
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
         Left            =   10560
         TabIndex        =   73
         Text            =   "Enter BHK"
         Top             =   6000
         Width           =   3855
      End
      Begin VB.CommandButton add_room_signup_command 
         Caption         =   "Add"
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
         Left            =   8400
         TabIndex        =   69
         Top             =   9360
         Width           =   6015
      End
      Begin VB.TextBox add_room_number_textbox 
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
         Left            =   10560
         TabIndex        =   68
         Text            =   "Enter room number"
         Top             =   4560
         Width           =   3855
      End
      Begin VB.Image Image22 
         Height          =   1050
         Left            =   5520
         Picture         =   "Rent Record.frx":10BFA
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label add_room_label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Room Amount"
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
         Left            =   8400
         TabIndex        =   72
         Top             =   5280
         Width           =   1260
      End
      Begin VB.Label add_room_label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "BHK"
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
         Left            =   8400
         TabIndex        =   71
         Top             =   6000
         Width           =   360
      End
      Begin VB.Label add_room_label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Services"
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
         Left            =   8400
         TabIndex        =   70
         Top             =   6720
         Width           =   705
      End
      Begin VB.Label add_room_label3 
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
         Left            =   8400
         TabIndex        =   67
         Top             =   4680
         Width           =   1305
      End
      Begin VB.Label add_room_label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Add Room Details"
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
         Left            =   10020
         TabIndex        =   66
         Top             =   3120
         Width           =   2610
      End
      Begin VB.Label add_room_label1 
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
         Left            =   720
         TabIndex        =   65
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label add_room_label7 
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
         Left            =   10980
         TabIndex        =   64
         Top             =   10320
         Width           =   885
      End
      Begin VB.Image Image3 
         Height          =   615
         Left            =   8400
         Picture         =   "Rent Record.frx":1A904
         Stretch         =   -1  'True
         Top             =   10200
         Width           =   6015
      End
      Begin VB.Image Image14 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   8160
         Picture         =   "Rent Record.frx":1AE0F
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   6495
      End
      Begin VB.Image Image2 
         Height          =   10455
         Left            =   6120
         Picture         =   "Rent Record.frx":1B319
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   10575
      End
   End
   Begin VB.Frame edit_room_frame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   12135
      Left            =   0
      TabIndex        =   86
      Top             =   0
      Width           =   23055
      Begin VB.Frame edit_room_frame_2 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   720
         TabIndex        =   93
         Top             =   4200
         Width           =   10215
         Begin VB.CheckBox edit_room_check_2 
            BackColor       =   &H8000000E&
            Caption         =   "Internet"
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
            Index           =   4
            Left            =   6240
            TabIndex        =   119
            Top             =   5160
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CheckBox edit_room_check_2 
            BackColor       =   &H8000000E&
            Caption         =   "Security"
            Enabled         =   0   'False
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
            Index           =   3
            Left            =   6240
            TabIndex        =   118
            Top             =   4680
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox edit_room_check_2 
            BackColor       =   &H8000000E&
            Caption         =   "Waste"
            Enabled         =   0   'False
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
            Index           =   2
            Left            =   6240
            TabIndex        =   117
            Top             =   4200
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox edit_room_check_2 
            BackColor       =   &H8000000E&
            Caption         =   "Drinking Water"
            Enabled         =   0   'False
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
            Left            =   6240
            TabIndex        =   116
            Top             =   3720
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.CheckBox edit_room_check_1 
            BackColor       =   &H8000000E&
            Caption         =   "Internet"
            Enabled         =   0   'False
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
            Index           =   4
            Left            =   2280
            TabIndex        =   115
            Top             =   5160
            Width           =   1935
         End
         Begin VB.CheckBox edit_room_check_1 
            BackColor       =   &H8000000E&
            Caption         =   "Security"
            Enabled         =   0   'False
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
            Index           =   3
            Left            =   2280
            TabIndex        =   114
            Top             =   4680
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox edit_room_check_1 
            BackColor       =   &H8000000E&
            Caption         =   "Waste"
            Enabled         =   0   'False
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
            Index           =   2
            Left            =   2280
            TabIndex        =   113
            Top             =   4200
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox edit_room_check_1 
            BackColor       =   &H8000000E&
            Caption         =   "Drinking Water"
            Enabled         =   0   'False
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
            Left            =   2280
            TabIndex        =   112
            Top             =   3720
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox edit_room_check_2 
            BackColor       =   &H8000000E&
            Caption         =   "Electricity"
            Enabled         =   0   'False
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
            Left            =   6240
            TabIndex        =   105
            Top             =   3240
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox edit_room_check_1 
            BackColor       =   &H8000000E&
            Caption         =   "Electricity"
            Enabled         =   0   'False
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
            Left            =   2280
            TabIndex        =   104
            Top             =   3240
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.TextBox edit_room_textbox 
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
            Index           =   0
            Left            =   6240
            TabIndex        =   96
            Text            =   "Enter room number"
            Top             =   1080
            Width           =   3615
         End
         Begin VB.TextBox edit_room_textbox 
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
            Index           =   1
            Left            =   6240
            TabIndex        =   95
            Text            =   "Enter rent amount"
            Top             =   1800
            Width           =   3615
         End
         Begin VB.TextBox edit_room_textbox 
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
            Index           =   2
            Left            =   6240
            TabIndex        =   94
            Text            =   "Enter BHK"
            Top             =   2520
            Width           =   3615
         End
         Begin VB.Shape Shape9 
            Height          =   5655
            Left            =   0
            Top             =   120
            Width           =   10215
         End
         Begin VB.Line Line9 
            BorderColor     =   &H80000010&
            X1              =   0
            X2              =   16320
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label edit_room_frame_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "New Detail"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   315
            Index           =   3
            Left            =   6240
            TabIndex        =   107
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label edit_room_frame_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Old Detail"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   315
            Index           =   2
            Left            =   2280
            TabIndex        =   106
            Top             =   240
            Width           =   1080
         End
         Begin VB.Line Line8 
            BorderColor     =   &H80000010&
            X1              =   6000
            X2              =   6000
            Y1              =   120
            Y2              =   5760
         End
         Begin VB.Line Line7 
            BorderColor     =   &H80000010&
            X1              =   2040
            X2              =   2040
            Y1              =   120
            Y2              =   5760
         End
         Begin VB.Label edit_room_frame_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "x"
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
            Index           =   12
            Left            =   2280
            TabIndex        =   103
            Top             =   2520
            Width           =   90
         End
         Begin VB.Label edit_room_frame_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "xxxxxx"
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
            Index           =   11
            Left            =   2280
            TabIndex        =   102
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label edit_room_frame_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "xxxx"
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
            Index           =   10
            Left            =   2280
            TabIndex        =   101
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label edit_room_frame_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Room Amount"
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
            Left            =   240
            TabIndex        =   100
            Top             =   1800
            Width           =   1260
         End
         Begin VB.Label edit_room_frame_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "BHK"
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
            Index           =   6
            Left            =   240
            TabIndex        =   99
            Top             =   2520
            Width           =   360
         End
         Begin VB.Label edit_room_frame_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Services"
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
            Index           =   7
            Left            =   240
            TabIndex        =   98
            Top             =   3240
            Width           =   705
         End
         Begin VB.Label edit_room_frame_label 
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
            Index           =   4
            Left            =   240
            TabIndex        =   97
            Top             =   1080
            Width           =   1305
         End
         Begin VB.Image Image6 
            Height          =   615
            Left            =   0
            Picture         =   "Rent Record.frx":21A65
            Stretch         =   -1  'True
            Top             =   120
            Width           =   10215
         End
      End
      Begin VB.Frame edit_room_frame_1 
         BackColor       =   &H8000000E&
         Height          =   2775
         Left            =   8993
         TabIndex        =   89
         Top             =   5400
         Width           =   4815
         Begin VB.CommandButton edit_room_1_command_1 
            Caption         =   "Proceed"
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
            Left            =   240
            TabIndex        =   92
            Top             =   1920
            Width           =   4335
         End
         Begin VB.ComboBox edit_room_1_combo_1 
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
            Left            =   1920
            TabIndex        =   90
            Text            =   "Combo1"
            Top             =   1080
            Width           =   2655
         End
         Begin VB.Label edit_room_1_label_1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Room Number"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   91
            Top             =   1080
            Width           =   1470
         End
         Begin VB.Label edit_room_1_label_2 
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Choose the room number from the list and press proceed for editing."
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   240
            TabIndex        =   108
            Top             =   240
            Width           =   4425
         End
      End
      Begin VB.Label edit_room_frame_save 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   154
         Top             =   10920
         Width           =   510
      End
      Begin VB.Label edit_room_frame_go_back 
         BackStyle       =   0  'Transparent
         Caption         =   "Go Back"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   1200
         TabIndex        =   153
         Top             =   10920
         Width           =   1215
      End
      Begin VB.Image Image23 
         Height          =   1050
         Left            =   5520
         Picture         =   "Rent Record.frx":21F6F
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Image edit_room_fr_image_1 
         Height          =   615
         Left            =   2760
         Picture         =   "Rent Record.frx":2BC79
         Stretch         =   -1  'True
         Top             =   10800
         Width           =   1935
      End
      Begin VB.Label edit_room_frame_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Edit Room"
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
         TabIndex        =   88
         Top             =   3000
         Width           =   1680
      End
      Begin VB.Image Image8 
         Height          =   615
         Left            =   720
         Picture         =   "Rent Record.frx":2C184
         Stretch         =   -1  'True
         Top             =   10800
         Width           =   1935
      End
      Begin VB.Image Image15 
         Appearance      =   0  'Flat
         Height          =   690
         Left            =   720
         Picture         =   "Rent Record.frx":2C68E
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   21585
      End
      Begin VB.Label edit_room_frame_label 
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
         TabIndex        =   87
         Top             =   480
         Width           =   4695
      End
   End
   Begin VB.Frame remove_room_frame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Remove Room Frame"
      ForeColor       =   &H80000008&
      Height          =   12255
      Left            =   0
      TabIndex        =   79
      Top             =   0
      Width           =   23055
      Begin VB.CommandButton remove_room_frame_remove_command 
         Caption         =   "Remove"
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
         Left            =   8873
         TabIndex        =   81
         Top             =   6840
         Width           =   5055
      End
      Begin VB.ComboBox remove_room_frame_combo1 
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
         Left            =   10913
         TabIndex        =   80
         Text            =   "Combo1"
         Top             =   6120
         Width           =   3015
      End
      Begin VB.Image Image24 
         Height          =   1050
         Left            =   5520
         Picture         =   "Rent Record.frx":2CB98
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label remove_room_frame_label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Available Rooms"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8880
         TabIndex        =   85
         Top             =   6120
         Width           =   1695
      End
      Begin VB.Label remove_room_frame_label5 
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
         TabIndex        =   84
         Top             =   7800
         Width           =   885
      End
      Begin VB.Image Image5 
         Height          =   615
         Left            =   8880
         Picture         =   "Rent Record.frx":368A2
         Stretch         =   -1  'True
         Top             =   7680
         Width           =   5055
      End
      Begin VB.Label remove_room_frame_label1 
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
         Left            =   720
         TabIndex        =   83
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label remove_room_frame_label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Remove Room"
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
         Left            =   10320
         TabIndex        =   82
         Top             =   4560
         Width           =   2130
      End
      Begin VB.Image Image16 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   8520
         Picture         =   "Rent Record.frx":36DAD
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   5655
      End
      Begin VB.Image Image4 
         Height          =   4575
         Left            =   6840
         Picture         =   "Rent Record.frx":372B7
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   9015
      End
   End
   Begin VB.Frame SignUpFrame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Sign Up Frame"
      Height          =   12375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   23055
      Begin VB.CheckBox signup_pw_hide 
         BackColor       =   &H8000000E&
         Height          =   435
         Left            =   8760
         TabIndex        =   162
         Top             =   5160
         Width           =   255
      End
      Begin VB.ComboBox SignUpFrame_dob_Combo 
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
         Index           =   2
         ItemData        =   "Rent Record.frx":3DA03
         Left            =   7680
         List            =   "Rent Record.frx":3DA67
         TabIndex        =   32
         Text            =   "Date"
         Top             =   7680
         Width           =   1095
      End
      Begin VB.ComboBox SignUpFrame_dob_Combo 
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
         Index           =   1
         ItemData        =   "Rent Record.frx":3DAE2
         Left            =   6000
         List            =   "Rent Record.frx":3DB0A
         TabIndex        =   31
         Text            =   "Month"
         Top             =   7680
         Width           =   1455
      End
      Begin VB.Timer SignUpFrame_Timer 
         Interval        =   2000
         Left            =   840
         Top             =   2880
      End
      Begin VB.CommandButton SignUpFrame_SignUp_Command 
         Caption         =   "Sign Up"
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
         Left            =   2640
         TabIndex        =   8
         Top             =   9570
         Width           =   6015
      End
      Begin VB.TextBox SignUpFrame_Password_TextBox 
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
         IMEMode         =   3  'DISABLE
         Left            =   4785
         TabIndex        =   7
         Text            =   "Enter password"
         Top             =   5250
         Width           =   3870
      End
      Begin VB.TextBox SignUpFrame_Username_TextBox 
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
         Left            =   4800
         TabIndex        =   6
         Text            =   "Enter username"
         Top             =   4530
         Width           =   3855
      End
      Begin VB.TextBox SignUpFrame_PasswordConfirmation_TextBox 
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
         IMEMode         =   3  'DISABLE
         Left            =   4800
         TabIndex        =   5
         Text            =   "Enter password for confirmation"
         Top             =   5880
         Width           =   3870
      End
      Begin VB.TextBox SignUpFrame_Pin_TextBox 
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
         IMEMode         =   3  'DISABLE
         Left            =   4800
         TabIndex        =   4
         Text            =   "Enter pin number"
         Top             =   7080
         Width           =   3870
      End
      Begin VB.TextBox SignUpFrame_Contact_TextBox 
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
         IMEMode         =   3  'DISABLE
         Left            =   4800
         MaxLength       =   25
         TabIndex        =   3
         Text            =   "Enter contact number"
         Top             =   6480
         Width           =   3870
      End
      Begin VB.TextBox SignUpFrame_Security_TextBox 
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
         IMEMode         =   3  'DISABLE
         Left            =   4800
         TabIndex        =   2
         Text            =   "Enter your favourite thing"
         Top             =   8280
         Width           =   3870
      End
      Begin VB.ComboBox SignUpFrame_dob_Combo 
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
         Index           =   0
         ItemData        =   "Rent Record.frx":3DB70
         Left            =   4800
         List            =   "Rent Record.frx":3DB72
         TabIndex        =   1
         Text            =   "Year"
         Top             =   7680
         Width           =   1095
      End
      Begin VB.Image Image19 
         Height          =   1050
         Left            =   5520
         Picture         =   "Rent Record.frx":3DB74
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label SignUpFrameLabel1 
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
         Left            =   720
         TabIndex        =   17
         Top             =   480
         Width           =   4695
      End
      Begin VB.Image SignUpFrameImage1 
         Height          =   12465
         Left            =   11760
         Picture         =   "Rent Record.frx":4787E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   13740
      End
      Begin VB.Label SignUpFrameLabel2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Sign Up"
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
         Left            =   5040
         TabIndex        =   16
         Top             =   3240
         Width           =   1125
      End
      Begin VB.Label SignUpFrameLabel4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   15
         Top             =   5280
         Width           =   930
      End
      Begin VB.Label SignUpFrameLabel3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   14
         Top             =   4530
         Width           =   990
      End
      Begin VB.Label SignUpFrameLabel7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Pin Number"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   13
         Top             =   7080
         Width           =   1170
      End
      Begin VB.Label SignUpFrameLabel6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Number"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   12
         Top             =   6480
         Width           =   1635
      End
      Begin VB.Label SignUpFrameLabel5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   11
         Top             =   5880
         Width           =   1785
      End
      Begin VB.Label SignUpFrameLabel9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Security Purpose"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   10
         Top             =   8280
         Width           =   1635
      End
      Begin VB.Label SignUpFrameLabel8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   9
         Top             =   7680
         Width           =   1275
      End
      Begin VB.Image Image11 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   2400
         Picture         =   "Rent Record.frx":51A12
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   6735
      End
      Begin VB.Image Image28 
         Height          =   9615
         Left            =   360
         Picture         =   "Rent Record.frx":51F1C
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   10815
      End
   End
   Begin VB.Frame ForgotPasswordFrame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Forgot Password Frame"
      Height          =   12375
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   23055
      Begin VB.ComboBox ForgotPasswordFrame_dob_Combo 
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
         Index           =   2
         ItemData        =   "Rent Record.frx":58668
         Left            =   7440
         List            =   "Rent Record.frx":586CC
         TabIndex        =   36
         Text            =   "Date"
         Top             =   8400
         Width           =   855
      End
      Begin VB.ComboBox ForgotPasswordFrame_dob_Combo 
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
         Index           =   1
         ItemData        =   "Rent Record.frx":58747
         Left            =   6000
         List            =   "Rent Record.frx":5876F
         TabIndex        =   35
         Text            =   "Month"
         Top             =   8400
         Width           =   1335
      End
      Begin VB.ComboBox ForgotPasswordFrame_dob_Combo 
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
         Index           =   0
         ItemData        =   "Rent Record.frx":587D5
         Left            =   4800
         List            =   "Rent Record.frx":5885A
         TabIndex        =   34
         Text            =   "Year"
         Top             =   8400
         Width           =   1095
      End
      Begin VB.TextBox ForgotPasswordFrame_Security_TextBox 
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
         IMEMode         =   3  'DISABLE
         Left            =   4800
         TabIndex        =   30
         Text            =   "Enter your favourite thing"
         Top             =   7680
         Width           =   3495
      End
      Begin VB.TextBox ForgotPasswordFrame_Pin_TextBox 
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
         IMEMode         =   3  'DISABLE
         Left            =   4800
         MaxLength       =   25
         TabIndex        =   29
         Text            =   "Enter pin number"
         Top             =   6960
         Width           =   3495
      End
      Begin VB.TextBox ForgotPasswordFrame_Username_TextBox 
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
         Left            =   4800
         TabIndex        =   21
         Text            =   "Enter username"
         Top             =   5520
         Width           =   3495
      End
      Begin VB.TextBox ForgotPasswordFrame_Contact_TextBox 
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
         IMEMode         =   3  'DISABLE
         Left            =   4800
         MaxLength       =   25
         TabIndex        =   20
         Text            =   "Enter contact number"
         Top             =   6240
         Width           =   3495
      End
      Begin VB.CommandButton ForgotPasswordFrame_Recover_Command 
         Caption         =   "Recover Password"
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
         Left            =   2880
         TabIndex        =   19
         Top             =   9120
         Width           =   5415
      End
      Begin VB.Timer ForgotPasswordFrame_Timer 
         Interval        =   2000
         Left            =   2760
         Top             =   1920
      End
      Begin VB.Image Image20 
         Height          =   1050
         Left            =   5520
         Picture         =   "Rent Record.frx":58960
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label ForgotPasswordFrameLabel2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Forgot Password"
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
         Left            =   4440
         TabIndex        =   23
         Top             =   4200
         Width           =   2445
      End
      Begin VB.Image Image12 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   2640
         Picture         =   "Rent Record.frx":6266A
         Stretch         =   -1  'True
         Top             =   4080
         Width           =   6015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Go back to Log In"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2880
         TabIndex        =   33
         Top             =   9960
         Width           =   1560
      End
      Begin VB.Label ForgotPasswordFrameLabel5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Pin Number"
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
         Left            =   3000
         TabIndex        =   28
         Top             =   6960
         Width           =   1035
      End
      Begin VB.Label ForgotPasswordFrameLabel6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Security purpose"
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
         Left            =   3000
         TabIndex        =   27
         Top             =   7680
         Width           =   1470
      End
      Begin VB.Label ForgotPasswordFrameLabel7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of birth"
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
         Left            =   3000
         TabIndex        =   26
         Top             =   8400
         Width           =   1110
      End
      Begin VB.Label ForgotPasswordFrameLabel3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
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
         Left            =   3000
         TabIndex        =   25
         Top             =   5520
         Width           =   885
      End
      Begin VB.Label ForgotPasswordFrameLabel4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Number"
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
         Left            =   3000
         TabIndex        =   24
         Top             =   6240
         Width           =   1440
      End
      Begin VB.Image ForgotPasswordFrame_Image1 
         Height          =   12465
         Left            =   11760
         Picture         =   "Rent Record.frx":62B74
         Stretch         =   -1  'True
         Top             =   0
         Width           =   13740
      End
      Begin VB.Label ForgotPasswordFrameLabel1 
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
         Left            =   720
         TabIndex        =   22
         Top             =   480
         Width           =   4695
      End
      Begin VB.Image Image29 
         Height          =   7815
         Left            =   840
         Picture         =   "Rent Record.frx":6CD08
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   9615
      End
   End
   Begin VB.Frame Reset_Frame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Log In Frame"
      Height          =   12375
      Left            =   0
      TabIndex        =   167
      Top             =   0
      Width           =   23055
      Begin VB.Timer reset_frame_timer 
         Interval        =   2000
         Left            =   960
         Top             =   3720
      End
      Begin VB.CommandButton reset_fr_reset 
         Caption         =   "Reset"
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
         Left            =   3000
         TabIndex        =   171
         Top             =   7530
         Width           =   5535
      End
      Begin VB.TextBox reset_fr_text_2 
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
         IMEMode         =   3  'DISABLE
         Left            =   5160
         MaxLength       =   31
         TabIndex        =   170
         Text            =   "Enter password for confirmation"
         Top             =   6720
         Width           =   3255
      End
      Begin VB.TextBox reset_fr_text_1 
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
         Left            =   5160
         TabIndex        =   169
         Text            =   "Enter password"
         Top             =   6000
         Width           =   3255
      End
      Begin VB.CheckBox reset_fr_label_check 
         BackColor       =   &H8000000E&
         Height          =   195
         Left            =   8520
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   168
         Top             =   6000
         Width           =   195
      End
      Begin VB.Label reset_fr_label_2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Reset"
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
         Left            =   5400
         TabIndex        =   176
         Top             =   4290
         Width           =   780
      End
      Begin VB.Label reset_fr_label_1 
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
         Left            =   720
         TabIndex        =   175
         Top             =   480
         Width           =   4695
      End
      Begin VB.Image reset_fr_image_1 
         Height          =   12465
         Left            =   11760
         Picture         =   "Rent Record.frx":73454
         Stretch         =   -1  'True
         Top             =   0
         Width           =   13740
      End
      Begin VB.Label reset_fr_label_5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3000
         TabIndex        =   174
         Top             =   8250
         Width           =   570
      End
      Begin VB.Label reset_fr_label_4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Password Confirmation"
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
         Left            =   3000
         TabIndex        =   173
         Top             =   6720
         Width           =   2025
      End
      Begin VB.Label reset_fr_label_3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   3000
         TabIndex        =   172
         Top             =   6000
         Width           =   840
      End
      Begin VB.Image reset_fr_image_2 
         Height          =   1050
         Left            =   5520
         Picture         =   "Rent Record.frx":7D5E8
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Image reset_fr_image_3 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   2760
         Picture         =   "Rent Record.frx":872F2
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   6135
      End
      Begin VB.Image reset_fr_image_4 
         Height          =   4695
         Left            =   960
         Picture         =   "Rent Record.frx":877FC
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   9735
      End
   End
   Begin VB.Frame LogInFrame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Log In Frame"
      Height          =   12375
      Left            =   0
      TabIndex        =   128
      Top             =   0
      Width           =   23055
      Begin VB.CheckBox login_pw_hide 
         BackColor       =   &H8000000E&
         Height          =   195
         Left            =   8280
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   161
         Top             =   6720
         Width           =   195
      End
      Begin VB.TextBox LogInFrame_Username_TextBox 
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
         Left            =   4920
         TabIndex        =   131
         Text            =   "Enter username"
         Top             =   6000
         Width           =   3255
      End
      Begin VB.TextBox LogInFrame_Password_TextBox 
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
         IMEMode         =   3  'DISABLE
         Left            =   4920
         MaxLength       =   25
         TabIndex        =   130
         Text            =   "Enter password"
         Top             =   6720
         Width           =   3255
      End
      Begin VB.CommandButton LogInFrameCommand1 
         Caption         =   "Log In"
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
         Left            =   3360
         TabIndex        =   129
         Top             =   7410
         Width           =   4815
      End
      Begin VB.Timer LogInFrame_Timer 
         Interval        =   2000
         Left            =   1080
         Top             =   1920
      End
      Begin VB.Label LogInFrameLabel6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Reset Password"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6720
         TabIndex        =   166
         Top             =   8160
         Width           =   1380
      End
      Begin VB.Image Image21 
         Height          =   1050
         Left            =   5520
         Picture         =   "Rent Record.frx":8DF48
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label LogInFrameLabel3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
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
         Left            =   3360
         TabIndex        =   136
         Top             =   6000
         Width           =   885
      End
      Begin VB.Label LogInFrameLabel4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   3360
         TabIndex        =   135
         Top             =   6720
         Width           =   840
      End
      Begin VB.Label LogInFrameLabel5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Forgot Password"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3360
         TabIndex        =   134
         Top             =   8130
         Width           =   1485
      End
      Begin VB.Image LogInFrameImage1 
         Height          =   12465
         Left            =   11760
         Picture         =   "Rent Record.frx":97C52
         Stretch         =   -1  'True
         Top             =   0
         Width           =   13740
      End
      Begin VB.Label LogInFrameLabel1 
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
         Left            =   720
         TabIndex        =   133
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label LogInFrameLabel2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Log In"
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
         Left            =   5520
         TabIndex        =   132
         Top             =   4290
         Width           =   900
      End
      Begin VB.Image Image13 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   3000
         Picture         =   "Rent Record.frx":A1DE6
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   5775
      End
      Begin VB.Image Image27 
         Height          =   4935
         Left            =   1320
         Picture         =   "Rent Record.frx":A22F0
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   9135
      End
   End
   Begin VB.Frame room_frame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12375
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   23055
      Begin VB.Frame Tenant_filter_frame 
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   17520
         TabIndex        =   155
         Top             =   10440
         Width           =   4695
         Begin VB.OptionButton room_filter_room_status 
            BackColor       =   &H80000016&
            Caption         =   "Unccupied"
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
            Index           =   1
            Left            =   480
            TabIndex        =   160
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton room_filter_room_status 
            BackColor       =   &H80000016&
            Caption         =   "Occupied"
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
            Index           =   0
            Left            =   480
            TabIndex        =   159
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox room_filter_option 
            BackColor       =   &H80000016&
            Caption         =   "Room status"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   157
            Top             =   120
            Width           =   1575
         End
         Begin VB.CommandButton room_filter_command 
            BackColor       =   &H8000000E&
            Caption         =   "Filter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2160
            TabIndex        =   156
            Top             =   600
            Width           =   2295
         End
         Begin VB.Shape Shape8 
            BorderColor     =   &H80000000&
            Height          =   1335
            Left            =   1920
            Top             =   0
            Width           =   2775
         End
         Begin VB.Shape Shape10 
            BorderColor     =   &H80000000&
            Height          =   1575
            Left            =   0
            Top             =   -240
            Width           =   1935
         End
         Begin VB.Label room_filter_status 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Filter status : Off"
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
            Left            =   2040
            TabIndex        =   158
            Top             =   240
            Width           =   2400
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   22200
         TabIndex        =   61
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   720
         TabIndex        =   59
         Top             =   3960
         Width           =   21495
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   720
         TabIndex        =   58
         Top             =   10080
         Width           =   21495
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   12960
         TabIndex        =   57
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   720
         TabIndex        =   56
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   720
         TabIndex        =   55
         Top             =   3480
         Width           =   21495
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   10800
         TabIndex        =   54
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   6600
         TabIndex        =   53
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   5280
         TabIndex        =   52
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000010&
         Height          =   6615
         Left            =   3600
         TabIndex        =   51
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   2880
         TabIndex        =   49
         Top             =   3480
         Width           =   15
         Begin VB.Frame Frame4 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   6735
            Left            =   120
            TabIndex        =   50
            Top             =   2520
            Width           =   135
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   1680
         TabIndex        =   48
         Top             =   3480
         Width           =   15
      End
      Begin VB.TextBox room_frame_TextBox 
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
         Height          =   5895
         Left            =   960
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   62
         Text            =   "Rent Record.frx":A8A3C
         Top             =   4080
         Width           =   21255
      End
      Begin VB.Image Image25 
         Height          =   1050
         Left            =   5520
         Picture         =   "Rent Record.frx":A8AD1
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label room_frame_label_11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Services Provided"
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
         Left            =   13080
         TabIndex        =   60
         Top             =   3600
         Width           =   2040
      End
      Begin VB.Label room_frame_label_7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Occupied"
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
         Left            =   5400
         TabIndex        =   47
         Top             =   3600
         Width           =   960
      End
      Begin VB.Label room_frame_label_5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "BHK"
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
         Left            =   3000
         TabIndex        =   46
         Top             =   3600
         Width           =   360
      End
      Begin VB.Label room_frame_label_4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Room No"
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
         Left            =   1800
         TabIndex        =   45
         Top             =   3600
         Width           =   840
      End
      Begin VB.Label room_frame_label_6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Rent Amount"
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
         Left            =   3720
         TabIndex        =   44
         Top             =   3600
         Width           =   1320
      End
      Begin VB.Label room_frame_label_10 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Tenant Contact"
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
         Left            =   10920
         TabIndex        =   43
         Top             =   3600
         Width           =   1680
      End
      Begin VB.Label room_frame_label_9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Tenant Name"
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
         Left            =   6720
         TabIndex        =   42
         Top             =   3600
         Width           =   1320
      End
      Begin VB.Label room_frame_main_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Go Back"
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
         Left            =   1080
         TabIndex        =   41
         Top             =   11160
         Width           =   885
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   600
         Picture         =   "Rent Record.frx":B27DB
         Stretch         =   -1  'True
         Top             =   11040
         Width           =   1935
      End
      Begin VB.Label room_frame_label_1 
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
         Left            =   720
         TabIndex        =   40
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label room_frame_label_3 
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
         Left            =   960
         TabIndex        =   39
         Top             =   3600
         Width           =   480
      End
      Begin VB.Label room_frame_label_2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Room Details"
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
         Left            =   1080
         TabIndex        =   38
         Top             =   2520
         Width           =   1920
      End
      Begin VB.Image Image17 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   720
         Picture         =   "Rent Record.frx":B2CE6
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Image Image26 
         Height          =   495
         Left            =   720
         Picture         =   "Rent Record.frx":B31F0
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   21495
      End
   End
   Begin VB.Shape Shape4 
      Height          =   495
      Left            =   10800
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   10800
      Top             =   5880
      Width           =   1215
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim count_1 As Integer
Dim count_2 As Integer
Dim count_3 As Integer

Dim counter_1 As Integer
Dim room_no As Integer
Dim option_count As Integer

Dim sn As Integer
Dim serial As String * 7
Dim service_1 As String * 57
Dim service_2 As String * 57

Dim MsgBox_Response As Integer
Dim LogInFrame_PictureTimer As Integer
Dim SignUpFrame_PictureTimer As Integer
Dim ForgotPasswordFrame_PictureTimer As Integer
Dim Reset_Frame_PictureTimer As Integer

Dim login As MainUser
Dim recover As MainUser
Dim file_data As MainUser

Dim flag As Boolean
Dim access As Boolean
Dim found As Boolean

Dim issue_1 As issue_class
Dim display_issue_1 As display_issue_class

Dim new_room As room_class
Dim temp_room As room_class
Dim existing_room As room_class

Dim tenant_1 As tenant_class
Dim tenant_2 As tenant_class
Dim tenant_3 As tenant_class

Dim service_class_1 As service
Dim service_class_2 As service
Dim service_class_3 As service

Dim display_service_class As display_service

Dim electricity_1 As electricity
Dim display_electricity_class As display_electricity

Dim room_display As add_room_display

Dim signup As MainUser
'Dim temp_room As room_class
Dim temp_issue As issue_class
Dim temp_tenant As tenant_class
Dim temp_service As service
Dim temp_electricity As electricity

'''''''''''''''''''''''''''''''''''''''Add Room Frame''''''''''''''''''''''''''''''''''''''''''''
'room number textbox
Private Sub add_room_number_textbox_GotFocus()
    If add_room_number_textbox.Text = "Enter room number" Then add_room_number_textbox.Text = ""
End Sub

Private Sub add_room_number_textbox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        KeyAscii = KeyAscii
    ElseIf KeyAscii = 13 Then
        If add_room_number_textbox <> "" Then
            KeyAscii = 0
            add_room_rent_textbox.SetFocus
        End If
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub add_room_number_textbox_LostFocus()
    If add_room_number_textbox = "" Then add_room_number_textbox.Text = "Enter room number"
End Sub

'rent amount textbox
Private Sub add_room_rent_textbox_GotFocus()
    If add_room_rent_textbox.Text = "Enter rent amount" Then add_room_rent_textbox.Text = ""
End Sub

Private Sub add_room_rent_textbox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 46 Then
        KeyAscii = KeyAscii
    ElseIf KeyAscii = 13 Then
        If add_room_rent_textbox <> "" Then
            KeyAscii = 0
            add_room_bhk_textbox.SetFocus
        End If
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub add_room_rent_textbox_LostFocus()
    If add_room_rent_textbox.Text = "" Then add_room_rent_textbox.Text = "Enter rent amount"
End Sub


'bhk textbox
Private Sub add_room_bhk_textbox_GotFocus()
    If add_room_bhk_textbox.Text = "Enter BHK" Then add_room_bhk_textbox.Text = ""
End Sub

Private Sub add_room_bhk_textbox_LostFocus()
    If add_room_bhk_textbox.Text = "" Then add_room_bhk_textbox.Text = "Enter BHK"
End Sub

Private Sub add_room_bhk_textbox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        KeyAscii = KeyAscii
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        If add_room_bhk_textbox.Text <> "" And add_room_bhk_textbox.Text <> "Enter BHK" Then
            add_room_signup_command_Click
        End If
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

'add room frame
Private Sub add_room_label7_Click()
    'reset user provided details
    add_room_number_textbox.Text = "Enter room number"
    add_room_rent_textbox.Text = "Enter rent amount"
    add_room_bhk_textbox.Text = "Enter BHK"
    'add_room_check1(4).Value = Unchecked
            
    'frame visibility
    main_menu_frame.Visible = True
    add_room_frame.Visible = False
End Sub

Private Sub add_room_signup_command_Click()
    'check for default values
    If add_room_number_textbox = "Enter room number" Or add_room_rent_textbox = "Enter rent amount" Or add_room_bhk_textbox = "Enter BHK" Then
        If add_room_number_textbox.Text = "Enter room number" Then
            MsgBox_Response = MsgBox("                    Please enter valid room number.", vbInformation + vbOKOnly, "Rental Record")
            add_room_number_textbox.SetFocus
        ElseIf add_room_rent_textbox.Text = "Enter rent amount" Then
            MsgBox_Response = MsgBox("                    Please enter valid rent amount.", vbInformation + vbOKOnly, "Rental Record")
            add_room_rent_textbox.SetFocus
        Else
            MsgBox_Response = MsgBox("                    Please enter valid BHK.", vbInformation + vbOKOnly, "Invalid Rental Record")
            add_room_bhk_textbox.SetFocus
        End If
        Exit Sub
    End If
    
    'check for existing room number for redundency
    count_1 = 1
    found = False
    temp_room.room_number = Val(add_room_number_textbox.Text)
    Open "roomdetail.txt" For Random As #1 Len = 112
        While EOF(1) <> True
            Get #1, count_1, existing_room
            count_1 = count_1 + 1
            If existing_room.room_number = temp_room.room_number Then found = True
        Wend
    Close #1
        
    If found = False Then
        'copy file dvalues to variables
        new_room.bhk = Val(add_room_bhk_textbox.Text)
        new_room.room_number = Val(add_room_number_textbox.Text)
        new_room.rent_amount = Val(add_room_rent_textbox.Text)
        
        new_room.room_occupied = False
        new_room.tenant_fname = "Null"
        new_room.tenant_mname = "Null"
        new_room.tenant_lname = "Null"
        new_room.tenant_contact = "Null"
        new_room.service_provided = "Null"
        
        'count number of existing detail
        room_detail_counT = room_detail_count_function()
        
        'write to the file
        Open "RoomDetail.txt" For Random As #2 Len = 112
            If room_detail_counT = 0 Then
                Put #2, 1, new_room
            Else
                Put #2, room_detail_counT + 1, new_room
            End If
        Close #2
        
        MsgBox_Response = MsgBox("     The details have been saved. Do you want to add another    details?", vbInformation + vbYesNo, "Room detail added successfully.")
        
        If MsgBox_Response = vbYes Then
            'reset textbox values
            add_room_number_textbox.Text = "Enter room number"
            add_room_rent_textbox.Text = "Enter rent amount"
            add_room_bhk_textbox.Text = "Enter BHK"
            
            add_room_number_textbox.SetFocus
        Else
            'reset textbox values
            add_room_number_textbox.Text = "Enter room number"
            add_room_rent_textbox.Text = "Enter rent amount"
            add_room_bhk_textbox.Text = "Enter BHK"
    
            'frame visibility
            main_menu_frame.Visible = True
            add_room_frame.Visible = False
        End If
    Else
        MsgBox_Response = MsgBox("                     Room with room number " & temp_room.room_number & " is already present.", vbInformation + vbOKOnly, "Rental Record")
    End If
End Sub


Private Sub edit_room_1_combo_1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub edit_room_1_command_1_Click()
    If edit_room_1_combo_1.Text <> "" And edit_room_1_combo_1.Text <> "Room number" Then
        'load existing room detail in edit room frame 2
        room_no = Val(edit_room_1_combo_1.Text)
        
        count_1 = room_detail_count_function
        
        Open "RoomDetail.txt" For Random As #1 Len = 112
            For count_2 = 1 To count_1 Step 1
                Get #1, count_2, existing_room
                If existing_room.room_number = room_no Then
                    edit_room_frame_label(10).Caption = existing_room.room_number
                    edit_room_frame_label(11).Caption = existing_room.rent_amount
                    edit_room_frame_label(12).Caption = existing_room.bhk
                    
                    'copying default values for new details for ease
                    edit_room_textbox(0).Text = Val(existing_room.room_number)
                    edit_room_textbox(1).Text = Val(existing_room.rent_amount)
                    edit_room_textbox(2).Text = Val(existing_room.bhk)
                    
                    'check values
                    edit_room_check_1(4).Enabled = True
                    
                    service_1 = "Water, Waste Management, Electricity, Security, Internet"
                    
                    If existing_room.room_occupied = True Then
                        If existing_room.service_provided = service_1 Then
                            edit_room_check_1(4).Value = Checked
                        Else
                            edit_room_check_1(4).Value = Unchecked
                        End If
                    Else
                        edit_room_check_1(4).Value = Unchecked
                    End If
                    
                    edit_room_check_1(4).Enabled = False
                End If
            Next count_2
        Close #1
        
        edit_room_frame_save.Visible = True
        edit_room_fr_image_1.Visible = True
        
        'frame visibility
        edit_room_frame_1.Visible = False
        edit_room_frame_2.Visible = True
    Else
        MsgBox_Response = MsgBox("               Please select the room first.", vbInformation + vbOKOnly, "Rental Recorded")
    End If
End Sub


Private Sub edit_room_frame_go_back_Click()
    If edit_room_frame_2.Visible = True Then
        'reset user provided values
        edit_room_textbox(0).Text = "Enter room number"
        edit_room_textbox(1).Text = "Enter rent amount"
        edit_room_textbox(2).Text = "Enter BHK"
        
        'frame visibility
        edit_room_frame_1.Visible = True
        edit_room_frame_2.Visible = False
        edit_room_frame_save.Visible = False
        edit_room_fr_image_1.Visible = False
        
        'reset combo box text
        edit_room_1_combo_1.Text = ""
    Else
        'reset user provided values
        edit_room_textbox(0).Text = "Enter room number"
        edit_room_textbox(1).Text = "Enter rent amount"
        edit_room_textbox(2).Text = "Enter BHK"
        
        edit_room_frame.Visible = False
        main_menu_frame.Visible = True
    End If
End Sub

'Edit room frame -> save command
Private Sub edit_room_frame_save_Click()
    'check for new values
    If edit_room_textbox(0).Text = "Enter room number" Or edit_room_textbox(1).Text = "Enter rent amount" Or edit_room_textbox(2).Text = "Enter BHK" Then
        If edit_room_textbox(0).Text = "Enter room number" Then
            MsgBox_Response = MsgBox("          Please enter the valid room number.", vbInformation + vbOKOnly, "Rental Record")
            edit_room_textbox(0).SetFocus
        ElseIf edit_room_textbox(1).Text = "Enter rent amount" Then
            MsgBox_Response = MsgBox("          Please enter the valid rent amount.", vbInformation + vbOKOnly, "Rental Reccord")
            edit_room_textbox(1).SetFocus
        ElseIf edit_room_textbox(2).Text = "Enter BHK" Then
            MsgBox_Response = MsgBox("          Please enter the valid BHK.", vbInformation + vbOKOnly, "Rental Record")
            edit_room_textbox(1).SetFocus
        End If
        Exit Sub
    End If
       
    If edit_room_check_2(4).Value = Unchecked Then
        new_room.service_provided = service_1
    ElseIf edit_room_check_2(4).Value = Checked Then
        new_room.service_provided = service_2
    End If
        
    'copy new values
    new_room.room_number = Val(edit_room_textbox(0))
    new_room.rent_amount = Val(edit_room_textbox(1))
    new_room.bhk = Val(edit_room_textbox(2))
        
    'setting default values
    new_room.room_occupied = False
    new_room.tenant_fname = "Null"
    new_room.tenant_mname = "Null"
    new_room.tenant_lname = "Null"
    new_room.tenant_contact = "Null"
        
    'check if room with same room name; check only if previous room number is changed
    found = False
    If new_room.room_number <> room_no Then
        Open "RoomDetail.txt" For Random As #2 Len = 112
            count_2 = room_detail_count_function
            For count_1 = 1 To count_2 Step 1
                Get #2, count_1, existing_room
                If existing_room.room_number = new_room.room_number Then found = True
            Next count_1
        Close #2
    End If
        
    If found = False Then 'no duplicated room found
        'write in a file
        count_1 = room_detail_count_function
        Open "RoomDetail.txt" For Random As #1 Len = 112
            For count_2 = 1 To count_1 Step 1
                Get #1, count_2, existing_room
                
                new_room.room_number = Val(edit_room_textbox(0).Text)
                new_room.rent_amount = Val(edit_room_textbox(1).Text)
                new_room.bhk = Val(edit_room_textbox(2).Text)
                
                new_room.room_occupied = existing_room.room_occupied
                new_room.tenant_fname = existing_room.tenant_fname
                new_room.tenant_mname = existing_room.tenant_mname
                new_room.tenant_lname = existing_room.tenant_lname
                new_room.tenant_contact = existing_room.tenant_contact
                
                If existing_room.room_number = room_no Then
                    service_1 = "Null"
                    If existing_room.service_provided = service_1 Then
                        new_room.service_provided = "Null"
                        Put #1, count_2, new_room
                    Else
                        Put #1, count_2, new_room
                    End If
                End If
            Next count_2
        Close #1
            
        'ask user if thay want to modify another room detail again
        MsgBox_Response = MsgBox("          Do you want to modify another room detail?", vbInformation + vbYesNo, "Room Detail Modified")
        
        If MsgBox_Response = 6 Then
            room_option_frame_edit_Click
        Else
            'reset user provided values
        
            'frame visibility
            edit_room_frame_save.Visible = False
            edit_room_frame.Visible = False
            main_menu_frame.Visible = True
        End If
    Else
        MsgBox_Response = MsgBox("          Please enter another room number as room with number " & new_room.room_number & " is already present.", vbInformation + vbOKOnly, "Rental Record")
        edit_room_textbox(0).SetFocus
    End If
End Sub

'edit room frame -> frame 2 -> Text Box
Private Sub edit_room_textbox_GotFocus(Index As Integer)
    If Index = 0 Then   'room number textbox
        If edit_room_textbox(0).Text = "Enter room number" Then edit_room_textbox(0).Text = ""
    ElseIf Index = 1 Then
        If edit_room_textbox(1).Text = "Enter rent amount" Then edit_room_textbox(1).Text = ""
    ElseIf Index = 2 Then
        If edit_room_textbox(2).Text = "Enter BHK" Then edit_room_textbox(2).Text = ""
    End If
End Sub

Private Sub edit_room_textbox_LostFocus(Index As Integer)
    If Index = 0 Then   'room number textbox
        If edit_room_textbox(0).Text = "" Then edit_room_textbox(0).Text = "Enter room number"
    ElseIf Index = 1 Then
        If edit_room_textbox(1).Text = "" Then edit_room_textbox(1).Text = "Enter rent amount"
    ElseIf Index = 2 Then
        If edit_room_textbox(2).Text = "" Then edit_room_textbox(2).Text = "Enter BHK"
    End If
End Sub

Private Sub edit_room_textbox_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 0 Then
            If edit_room_textbox(0).Text <> "" Then edit_room_textbox(1).SetFocus
        ElseIf Index = 1 Then
            If edit_room_textbox(1).Text <> "" Then edit_room_textbox(2).SetFocus
        End If
    ElseIf KeyAscii = 8 Then
        KeyAscii = KeyAscii
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub ForgotPasswordFrame_dob_Combo_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        If KeyAscii = 8 Then
            KeyAscii = KeyAscii
        ElseIf KeyAscii = 13 And ForgotPasswordFrame_dob_Combo(0).Text <> "" Then
            ForgotPasswordFrame_dob_Combo(1).SetFocus
        ElseIf KeyAscii < 47 Or KeyAscii > 57 Then
            KeyAscii = 0
        End If
    ElseIf Index = 1 Then
        If KeyAscii = 13 And ForgotPasswordFrame_dob_Combo(1).Text <> "Month" Then
            KeyAscii = 0
            ForgotPasswordFrame_dob_Combo(2).SetFocus
        Else
            KeyAscii = 0
        End If
    ElseIf Index = 2 Then
        If keyacii = 13 And ForgotPasswordFrame_dob_Combo(2).Text <> "Date" Then
            KeyAscii = 0
        Else
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub ForgotPasswordFrame_dob_Combo_LostFocus(Index As Integer)
    If Index = 0 Then
        If ForgotPasswordFrame_dob_Combo(0).Text = "" Then ForgotPasswordFrame_dob_Combo(0).Text = "Year"
    End If
End Sub

Private Sub ForgotPasswordFrame_Recover_Command_Click()
    'check fo default values
    If ForgotPasswordFrame_Username_TextBox.Text = "Enter username" Or ForgotPasswordFrame_Contact_TextBox.Text = "Enter conatct number" Or ForgotPasswordFrame_Pin_TextBox.Text = "Enter pin number" Or ForgotPasswordFrame_Security_TextBox.Text = "Enter your favourite thing" Then
        If ForgotPasswordFrame_Username_TextBox.Text = "Enter username" Then
            MsgBox_Response = MsgBox("                    Please enter the valid username.", vbInformation + vbOKOnly, "Rental Record")
            ForgotPasswordFrame_Username_TextBox.SetFocus
            
        ElseIf ForgotPasswordFrame_Contact_TextBox.Text = "Enter contact number" Then
            MsgBox_Response = MsgBox("                    Please enter the valid contact number.", vbInformation + vbOKOnly, "Rental Record")
            ForgotPasswordFrame_Contact_TextBox.SetFocus
            
        ElseIf ForgotPasswordFrame_Pin_TextBox.Text = "Enter pin number" Then
            MsgBox_Response = MsgBox("                    Please enter the valid pin number.", vbInformation + vbOKOnly, "Rental Record")
            ForgotPasswordFrame_Pin_TextBox.SetFocus
            
        ElseIf ForgotPasswordFrame_Security_TextBox.Text = "Enter your favourite thing" Then
            MsgBox_Response = MsgBox("                    Please answer the seurity question.", vbInformation + vbOKOnly, "Rental Record")
            ForgotPasswordFrame_Security_TextBox.SetFocus
        End If
        
        Exit Sub
    End If
    
    'check the validity for combo box
    If ForgotPasswordFrame_dob_Combo(0).Text = "Year" Or ForgotPasswordFrame_dob_Combo(1).Text = "Month" Or ForgotPasswordFrame_dob_Combo(2).Text = "Date" Then
        If ForgotPasswordFrame_dob_Combo(0).Text = "Year" Then
            MsgBox_Response = MsgBox("                    Please choose the valid year.", vbInformation + vbOKOnly, "Rental Record")
            ForgotPasswordFrame_dob_Combo(0).SetFocus
        
        ElseIf ForgotPasswordFrame_dob_Combo(1).Text = "Month" Then
            MsgBox_Response = MsgBox("                    Please choose the valid month.", vbInformation + vbOKOnly, "Rental Record")
            ForgotPasswordFrame_dob_Combo(1).SetFocus
            
        ElseIf ForgotPasswordFrame_dob_Combo(2).Text = "Date" Then
            MsgBox_Response = MsgBox("                    Please choose the valid date.", vbInformation + vbOKOnly, "Rental Record")
            ForgotPasswordFrame_dob_Combo(2).SetFocus
        End If
        
        Exit Sub
    End If
    
    Open "admindata.txt" For Random As #1 Len = 121
        Get #1, 1, file_data
    Close #1
    
    recover.username = ForgotPasswordFrame_Username_TextBox.Text
    recover.contact_number = ForgotPasswordFrame_Contact_TextBox.Text
    recover.pin_number = ForgotPasswordFrame_Pin_TextBox.Text
    recover.security_question = ForgotPasswordFrame_Security_TextBox.Text
    recover.dob_year = ForgotPasswordFrame_dob_Combo(0).Text
    recover.dob_month = ForgotPasswordFrame_dob_Combo(1).Text
    recover.dob_date = ForgotPasswordFrame_dob_Combo(2).Text
    
    'data validation for showing password
    access = False
    If recover.username = file_data.username Then
        If recover.contact_number = file_data.contact_number Then
            If recover.pin_number = file_data.pin_number Then
                If recover.security_question = file_data.security_question Then
                    If recover.dob_year = file_data.dob_year Then
                        If recover.dob_month = file_data.dob_month Then
                            If recover.dob_date = file_data.dob_date Then
                                access = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If access = True Then
        MsgBox_Response = MsgBox("               Your password is : " & file_data.password, vbInformation + vbOKOnly, "Rental Record")
        LogInFrame.Visible = True
        ForgotPasswordFrame.Visible = False
    Else
        MsgBox_Response = MsgBox("      Make sure you entered the details correctly.", vbInformation + vbOKOnly, "Rental Record")
    End If
End Sub

Private Sub ForgotPasswordFrame_Timer_Timer()
    If ForgotPasswordFrame_PictureTimer < 4 Then
        ForgotPasswordFrame_PictureTimer = ForgotPasswordFrame_PictureTimer + 1
    Else
        ForgotPasswordFrame_PictureTimer = 1
    End If
    
    If ForgotPasswordFrame_PictureTimer = 1 Then
        ForgotPasswordFrame_Image1.Picture = LoadPicture("Assests/ApartmentPicture-1.jpg")
    ElseIf ForgotPasswordFrame_PictureTimer = 2 Then
        ForgotPasswordFrame_Image1.Picture = LoadPicture("Assests/ApartmentPicture-2.jpg")
    ElseIf ForgotPasswordFrame_PictureTimer = 3 Then
        ForgotPasswordFrame_Image1.Picture = LoadPicture("Assests/ApartmentPicture-3.jpg")
    Else
        ForgotPasswordFrame_Image1.Picture = LoadPicture("Assests/ApartmentPicture-4.jpg")
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Forgot Password Frame'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'username textbox
Private Sub ForgotPasswordFrame_Username_TextBox_GotFocus()
    If ForgotPasswordFrame_Username_TextBox.Text = "Enter username" Then ForgotPasswordFrame_Username_TextBox.Text = ""
End Sub

Private Sub ForgotPasswordFrame_Username_TextBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 And ForgotPasswordFrame_Username_TextBox <> "" Then
        KeyAscii = 0
        ForgotPasswordFrame_Contact_TextBox.SetFocus
    End If
End Sub

Private Sub ForgotPasswordFrame_Username_TextBox_LostFocus()
    If ForgotPasswordFrame_Username_TextBox.Text = "" Then ForgotPasswordFrame_Username_TextBox.Text = "Enter username"
End Sub


'contact number textbox
Private Sub ForgotPasswordFrame_Contact_TextBox_GotFocus()
    If ForgotPasswordFrame_Contact_TextBox.Text = "Enter contact number" Then ForgotPasswordFrame_Contact_TextBox.Text = ""
End Sub

Private Sub ForgotPasswordFrame_contact_TextBox_LostFocus()
    If ForgotPasswordFrame_Contact_TextBox.Text = "" Then ForgotPasswordFrame_Contact_TextBox.Text = "Enter contact number"
End Sub

Private Sub ForgotPasswordFrame_contact_TextBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        KeyAscii = KeyAscii
    ElseIf KeyAscii = 13 And ForgotPasswordFrame_Contact_TextBox.Text <> "" Then
        KeyAscii = 0
        ForgotPasswordFrame_Pin_TextBox.SetFocus
    ElseIf KeyAscii < 47 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

'pin number textbox
Private Sub ForgotPasswordFrame_pin_TextBox_GotFocus()
    If ForgotPasswordFrame_Pin_TextBox.Text = "Enter pin number" Then ForgotPasswordFrame_Pin_TextBox.Text = ""
End Sub

Private Sub ForgotPasswordFrame_pin_TextBox_LostFocus()
    If ForgotPasswordFrame_Pin_TextBox.Text = "" Then ForgotPasswordFrame_Pin_TextBox.Text = "Enter pin number"
End Sub

Private Sub ForgotPasswordFrame_Pin_TextBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        KeyAscii = KeyAscii
    ElseIf KeyAscii = 13 And ForgotPasswordFrame_Pin_TextBox <> "" Then
        KeyAscii = 0
        ForgotPasswordFrame_Security_TextBox.SetFocus
    ElseIf KeyAscii > 47 And KeyAscii < 57 Then
        text_length = Len(ForgotPasswordFrame_Pin_TextBox.Text)
        If text_length > 3 Then KeyAscii = 0
    ElseIf KeyAscii < 47 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Tenant_Form
    Unload Issue_Form
    Unload Payment_Form
End Sub

Private Sub issue_option_details_Click()
    'load issues in a textbox
    Issue_Form.issue_detail_TextBox = ""
    Issue_Form.issue_filter_filter_status.Caption = "Filter : Off "
    Issue_Form.issue_filter_option(0).Value = True
    
    count_2 = issue_count_function
    Open "IssueDetail.txt" For Random As #1 Len = 167
        If count_2 > 0 Then
            For count_1 = 1 To count_2 Step 1
                Get #1, count_1, issue_1
                display_issue_1.serial = CStr(count_1)
                display_issue_1.room_num = CStr(issue_1.room_num)
                display_issue_1.reporter = issue_1.reporter
                display_issue_1.contact_num = issue_1.contact_num
                display_issue_1.issue_status = issue_1.issue_status
                display_issue_1.issue_reported_date = Format(issue_1.issue_reported_date, "dd mmmm, yyyy")
                
                If issue_1.issue_solved_date = 0 Then
                    display_issue_1.issue_solved_date = "-"
                Else
                    display_issue_1.issue_solved_date = Format(issue_1.issue_solved_date, "dd mmmm, yyyy")
                End If
                
                'display
                Issue_Form.issue_detail_TextBox = Issue_Form.issue_detail_TextBox + display_issue_1.serial + display_issue_1.room_num + display_issue_1.reporter
                Issue_Form.issue_detail_TextBox = Issue_Form.issue_detail_TextBox + display_issue_1.contact_num + display_issue_1.issue_status
                Issue_Form.issue_detail_TextBox = Issue_Form.issue_detail_TextBox + display_issue_1.issue_reported_date + display_issue_1.issue_solved_date
                Issue_Form.issue_detail_TextBox = Issue_Form.issue_detail_TextBox + issue_1.issue_detail + vbNewLine
            Next count_1
        Else 'has no issues at present
            Issue_Form.issue_detail_TextBox.Text = "No issues recorded at present!"
        End If
    Close #1
    
    'form & frame visibility
    Issue_Form.WindowState = Main_Form.WindowState
    Issue_Form.Visible = True
    Issue_Form.issue_detail_frame.Visible = True
    Issue_Form.issue_report_frame.Visible = False
    Main_Form.Visible = False
End Sub

Private Sub issue_option_details_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    issue_option_image(0).Visible = True
    issue_option_image(1).Visible = False
    
    main_manu_note_label.Caption = "Note : Click here to view issue details."
End Sub

Private Sub issue_option_report_Click()
    'load room number in combo box of issue form
    count_2 = room_detail_count_function
    Issue_Form.issue_fr_roomNum_combo.Clear
    Open "RoomDetail.txt" For Random As #1 Len = 112
        For count_1 = 1 To count_2 Step 1
            Get #1, count_1, existing_room
            Issue_Form.issue_fr_roomNum_combo.AddItem existing_room.room_number
        Next count_1
    Close #1
    
    'reset textbox values
    Issue_Form.issue_fr_roomNum_combo.Text = "Room number"
    Issue_Form.issue_fr_reporter_textbox.Text = "Enter reporter's name"
    Issue_Form.issue_fr_contact_textbox.Text = "Enter contact number"
    Issue_Form.issue_fr_issue_textbox.Text = ""
    
    'form and frame visibility
    Issue_Form.WindowState = Main_Form.WindowState
    Issue_Form.Visible = True
    Issue_Form.issue_report_frame.Visible = True
    Issue_Form.issue_detail_frame.Visible = False
    Main_Form.Visible = False
End Sub

Private Sub issue_option_report_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    issue_option_image(0).Visible = False
    issue_option_image(1).Visible = True
    main_manu_note_label.Caption = "Note : Click here to report a new issue."
End Sub

'Forgot password >> Back to lo log In
Private Sub Label1_Click()
    'reset user provided details
    ForgotPasswordFrame_Username_TextBox.Text = "Enter username"
    ForgotPasswordFrame_Contact_TextBox.Text = "Enter contact number"
    ForgotPasswordFrame_Pin_TextBox.Text = "Enter pin number"
    ForgotPasswordFrame_Security_TextBox.Text = "Enter your favourite thing"
    
    ForgotPasswordFrame_dob_Combo(0).Text = "Year"
    ForgotPasswordFrame_dob_Combo(1).Text = "Month"
    ForgotPasswordFrame_dob_Combo(2).Text = "Date"
    
    'frame visibility
    LogInFrame.Visible = True
    SignUpFrame.Visible = False
    ForgotPasswordFrame.Visible = False
    
    'timer set and reset
    LogInFrame_Timer.Enabled = True
    SignUpFrame_Timer.Enabled = False
    ForgotPasswordFrame_Timer.Enabled = False
End Sub

Private Sub login_pw_hide_Click()
    If login_pw_hide.Value = Checked Then
        LogInFrame_Password_TextBox.PasswordChar = ""
    Else
        LogInFrame_Password_TextBox.PasswordChar = "*"
    End If
End Sub

Private Sub LogInFrameCommand1_Click()
    LogInFrame_Timer.Enabled = False
    SignUpFrame_Timer.Enabled = False
    ForgotPasswordFrame_Timer.Enabled = False
        
    'check for emtiness of textboxes
    If LogInFrame_Username_TextBox.Text = "" Or LogInFrame_Password_TextBox.Text = "" Then
        If LogInFrame_Username_TextBox.Text = "" Then
            MsgboxResponse = MsgBox("                         Please enter the username.", vbInformation + vbOKOnly, "Rental Record")
        Else
            MsgboxResponse = MsgBox("                         Please enter the password.", vbInformation + vbOKOnly, "Rental Record")
        End If
        Exit Sub
    End If
    
    'check for default values
    If LogInFrame_Username_TextBox.Text = "Enter username" Or LogInFrame_Password_TextBox.Text = "Enter password" Then
        If LogInFrame_Username_TextBox.Text = "Enter username" Then
            MsgboxResponse = MsgBox("                         Please enter the valid username.", vbInformation + vbOKOnly, "Rental Record")
        Else
            MsgboxResponse = MsgBox("                         Please enter the valid password.", vbInformation + vbOKOnly, "Rental Record")
        End If
        Exit Sub
    End If
    
    'not empty textbox & no default values encountered
    'check for validity for logging in
    
    Open "AdminData.txt" For Random As #1 Len = 121
        Get #1, 1, signup
    Close #1
    
    login.username = LogInFrame_Username_TextBox.Text
    login.password = LogInFrame_Password_TextBox.Text
    
    If login.username = signup.username And login.password = signup.password Then
        'reset textbox values
        LogInFrame_Username_TextBox.Text = "Enter username"
        LogInFrame_Password_TextBox.Text = "Enter password"
        LogInFrame_Password_TextBox.PasswordChar = ""
        login_pw_hide.Value = Checked
        
        'frame visibility
        LogInFrame.Visible = False
        main_menu_frame.Visible = True
        room_option_frame.Visible = False
        tenant_option_frame.Visible = False
        payment_option_frame.Visible = False
        issue_option_frame.Visible = False
    Else
        MsgBox_Response = MsgBox("     Please make user you entered your username and password correctly.", vbInformation + vbOKOnly, "Rental Record")
    End If
End Sub

Private Sub LogInFrameLabel5_Click()
    'reset textbox values
    LogInFrame_Username_TextBox.Text = "Enter username"
    LogInFrame_Password_TextBox.Text = "Enter password"
    
    'frame visibility
    LogInFrame.Visible = False
    SignUpFrame.Visible = False
    ForgotPasswordFrame.Visible = True
    
    ForgotPasswordFrame_Timer.Enabled = True
    LogInFrame_Timer.Enabled = False
    ForgotPasswordFrame_PictureTimer = 1
End Sub

Private Sub LogInFrame_Username_TextBox_GotFocus()
    If LogInFrame_Username_TextBox.Text = "Enter username" Then LogInFrame_Username_TextBox.Text = ""
End Sub

Private Sub LogInFrame_Username_TextBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And LogInFrame_Username_TextBox.Text <> "" Then
        KeyAscii = 0
        LogInFrame_Password_TextBox.SetFocus
    ElseIf KeyAscii = 13 And LogInFrame_Username_TextBox.Text = "" Then
        KeyAscii = 0
    End If
End Sub

Private Sub LogInFrame_Username_TextBox_LostFocus()
    If LogInFrame_Username_TextBox.Text = "" Then LogInFrame_Username_TextBox.Text = "Enter username"
End Sub

Private Sub LogInFrame_Password_TextBox_GotFocus()
    If LogInFrame_Password_TextBox.Text = "Enter password" Then
        LogInFrame_Password_TextBox.PasswordChar = "*"
        login_pw_hide.Value = Unchecked
        LogInFrame_Password_TextBox.Text = ""
    End If
End Sub

Private Sub LogInFrame_Password_TextBox_LostFocus()
    If LogInFrame_Password_TextBox.Text = "" Then
        LogInFrame_Password_TextBox.Text = "Enter password"
        LogInFrame_Password_TextBox.PasswordChar = ""
        login_pw_hide.Value = Checked
    End If
End Sub

Private Sub LogInFrame_Password_TextBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And LogInFrame_Password_TextBox.Text <> "" Then
        KeyAscii = 0
        LogInFrameCommand1_Click
    End If
End Sub


Private Sub LogInFrame_Timer_Timer()
    If LogInFrame_PictureTimer < 4 Then
        LogInFrame_PictureTimer = LogInFrame_PictureTimer + 1
    Else
        LogInFrame_PictureTimer = 1
    End If
    
    If LogInFrame_PictureTimer = 1 Then
        LogInFrameImage1.Picture = LoadPicture("Assests/ApartmentPicture-1.jpg")
    ElseIf LogInFrame_PictureTimer = 2 Then
        LogInFrameImage1.Picture = LoadPicture("Assests/ApartmentPicture-2.jpg")
    ElseIf LogInFrame_PictureTimer = 3 Then
        LogInFrameImage1.Picture = LoadPicture("Assests/ApartmentPicture-3.jpg")
    Else
        LogInFrameImage1.Picture = LoadPicture("Assests/ApartmentPicture-4.jpg")
    End If
End Sub


'reset password
Private Sub LogInFrameLabel6_Click()
    'check for emtiness of textboxes
    If LogInFrame_Username_TextBox.Text = "" Or LogInFrame_Password_TextBox.Text = "" Then
        If LogInFrame_Username_TextBox.Text = "" Then
            MsgboxResponse = MsgBox("                         Please enter the username.", vbInformation + vbOKOnly, "Rental Record")
        Else
            MsgboxResponse = MsgBox("                         Please enter the password.", vbInformation + vbOKOnly, "Rental Record")
        End If
        Exit Sub
    End If
    
    'check for default values
    If LogInFrame_Username_TextBox.Text = "Enter username" Or LogInFrame_Password_TextBox.Text = "Enter password" Then
        If LogInFrame_Username_TextBox.Text = "Enter username" Then
            MsgboxResponse = MsgBox("                         Please enter the valid username.", vbInformation + vbOKOnly, "Rental Record")
        Else
            MsgboxResponse = MsgBox("                         Please enter the valid password.", vbInformation + vbOKOnly, "Rental Record")
        End If
        Exit Sub
    End If
    
    'check for validity for logging in
    Open "AdminData.txt" For Random As #1 Len = 121
        Get #1, 1, signup
    Close #1
    
    login.username = LogInFrame_Username_TextBox.Text
    login.password = LogInFrame_Password_TextBox.Text
    
    If login.username = signup.username And login.password = signup.password Then
        'reset values
        reset_fr_label_check.Value = Unchecked
        
        LogInFrame_Timer.Enabled = False
        reset_frame_timer.Enabled = True
    
        'frame visibility
        Reset_Frame.Visible = True
        LogInFrame.Visible = False
    Else
        MsgBox_Response = MsgBox("  Please make user you entered your username and password correctly.", vbInformation + vbOKOnly, "Rental Record")
    End If
End Sub


Private Sub main_menu_issue_option_Click()
    issue_option_details_Click
End Sub


Private Sub main_menu_logout_option_Click()
    MsgBox_Response = MsgBox("     Are you sure you want to log out?", vbInformation + vbYesNo, "Rental Record")
    If MsgBox_Response = vbYes Then
        'frame visibility
        main_menu_frame.Visible = False
        LogInFrame.Visible = True
    Else
        main_menu_frame.Visible = True
    End If
End Sub


Private Sub main_menu_payment_option_Click()
    payment_option_report_Click
End Sub

'''''''''''''''''''''''''''''''''''''
'''''main menu'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''
'Main menu -> Option room
Private Sub main_menu_room_option_Click()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    count_1 = 1
    serial = 1
    room_filter_room_status(0).Value = True
    
    room_filter_status.Caption = "Filter Status : Off"
    
    count_2 = room_detail_count_function
        
    If count_2 = 0 Then
        room_frame_TextBox.Text = "No detail has been added yet!"
    Else
        room_frame_TextBox.Text = ""
        Open "RoomDetail.txt" For Random As #1 Len = 112
            For count_1 = 1 To count_2 Step 1
                Get #1, count_1, existing_room
                room_display.sn = CStr(count_1)
                room_display.room_num = CStr(existing_room.room_number)
                room_display.bhk = CStr(existing_room.bhk)
                room_display.rent = CStr(existing_room.rent_amount)
                room_display.occupied = CStr(existing_room.room_occupied)
                    
                If existing_room.room_occupied = True Then
                    If existing_room.tenant_mname <> "Null      " Then 'has middle name
                        room_display.tenant_fname = existing_room.tenant_fname
                        room_display.tenant_mname = existing_room.tenant_mname
                        room_display.tenant_lname = existing_room.tenant_lname
                        room_display.tenant_contact = existing_room.tenant_contact
                    Else
                        room_display.tenant_fname = existing_room.tenant_fname
                        room_display.tenant_mname = existing_room.tenant_lname
                        room_display.tenant_lname = ""
                        room_display.tenant_contact = existing_room.tenant_contact
                    End If
                Else
                    room_display.tenant_fname = "-"
                    room_display.tenant_mname = "-"
                    room_display.tenant_lname = "-"
                    room_display.tenant_contact = "-"
                End If
                    
                'for display purpose only
                room_frame_TextBox.Text = room_frame_TextBox.Text + room_display.sn + room_display.room_num + room_display.bhk + room_display.rent + room_display.occupied
                room_frame_TextBox.Text = room_frame_TextBox.Text + room_display.tenant_fname + room_display.tenant_mname + room_display.tenant_lname
                room_frame_TextBox.Text = room_frame_TextBox.Text + room_display.tenant_contact + existing_room.service_provided + vbNewLine
                serial = serial + 1
            Next count_1
        Close #1
    End If
    
    'frame visibility
    LogInFrame.Visible = False
    SignUpFrame.Visible = False
    ForgotPasswordFrame.Visible = False
    main_menu_frame.Visible = False
    room_frame.Visible = True
    add_room_frame.Visible = False
End Sub

Private Sub main_menu_room_option_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    option_image(0).Visible = True
    option_image(1).Visible = False
    option_image(2).Visible = False
    option_image(3).Visible = False
    option_image(4).Visible = False
    
    room_option_frame.Visible = True
    tenant_option_frame.Visible = False
    payment_option_frame.Visible = False
    issue_option_frame.Visible = False
    
    room_option_image(0).Visible = True
    room_option_image(1).Visible = False
    room_option_image(2).Visible = False
    room_option_image(3).Visible = False
    
    main_manu_note_label.Caption = "Note : Click here to see the room details, add new details and remove or modify the details."
End Sub

Private Sub main_menu_tenant_option_Click()
    tenant_option_frame_view_Click
End Sub

Private Sub main_menu_tenant_option_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    option_image(0).Visible = False
    option_image(1).Visible = True
    option_image(2).Visible = False
    option_image(3).Visible = False
    option_image(4).Visible = False
    
    room_option_frame.Visible = False
    tenant_option_frame.Visible = True
    payment_option_frame.Visible = False
    issue_option_frame.Visible = False
    
    tenant_option_image(0).Visible = True
    tenant_option_image(1).Visible = False
    tenant_option_image(2).Visible = False
    tenant_option_image(3).Visible = False
    
    main_manu_note_label.Caption = "Note : Click here to see the tenant details, add new details and remove or modify the details."
End Sub

Private Sub main_menu_payment_option_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    option_image(0).Visible = False
    option_image(1).Visible = False
    option_image(2).Visible = True
    option_image(3).Visible = False
    option_image(4).Visible = False
    
    room_option_frame.Visible = False
    tenant_option_frame.Visible = False
    payment_option_frame.Visible = True
    issue_option_frame.Visible = False
    
    payment_option_image(0).Visible = True
    payment_option_image(1).Visible = False
    
    payment_option_image(0).Visible = True
    payment_option_image(1).Visible = False
    payment_option_image(2).Visible = False
    payment_option_image(3).Visible = False
    payment_option_image(4).Visible = False
    
    main_manu_note_label.Caption = "Note : Click here to see the payment report add do new payment."
End Sub

Private Sub main_menu_issue_option_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    option_image(0).Visible = False
    option_image(1).Visible = False
    option_image(2).Visible = False
    option_image(3).Visible = True
    option_image(4).Visible = False
    
    room_option_frame.Visible = False
    tenant_option_frame.Visible = False
    payment_option_frame.Visible = False
    issue_option_frame.Visible = True
    
    issue_option_image(0).Visible = True
    issue_option_image(1).Visible = False
    
    main_manu_note_label.Caption = "Note : Click here to see the reported issues and solve them."
End Sub

Private Sub main_menu_logout_option_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    option_image(0).Visible = False
    option_image(1).Visible = False
    option_image(2).Visible = False
    option_image(3).Visible = False
    option_image(4).Visible = True
    
    room_option_frame.Visible = False
    tenant_option_frame.Visible = False
    payment_option_frame.Visible = False
    issue_option_frame.Visible = False
    
    main_manu_note_label.Caption = "Note : Click here to log out."
End Sub

Private Sub main_menu_frame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.Caption = X
    Label3.Caption = Y
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''Main Menu Frame''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub option_frame_payment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    payment_option_image(0).Visible = False
    payment_option_image(1).Visible = True
End Sub

Private Sub option_frame_show_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    payment_option_image(0).Visible = True
    payment_option_image(1).Visible = False
End Sub


Private Sub payment_option_new_electricity_fee_Click()
    'form and frame visibility
    Payment_Form.WindowState = Main_Form.WindowState
    Payment_Form.Visible = True
    
    Payment_Form.unit_combo(0).Clear
    For count_1 = Format(Now, "yyyy") To Format(Now, "yyyy") + 1 Step 1
        Payment_Form.unit_combo(0).AddItem count_1
    Next count_1
    Payment_Form.unit_combo(0).Text = "Year"
    
    Main_Form.Visible = False
    Payment_Form.electricity_fee_frame.Visible = True
    Payment_Form.service_det_frame.Visible = False
    Payment_Form.service_add_fr.Visible = False
    Payment_Form.payment_frame.Visible = False
    Payment_Form.payment_detail_frame.Visible = False
End Sub

Private Sub payment_option_new_electricity_fee_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    payment_option_image(0).Visible = False
    payment_option_image(1).Visible = False
    payment_option_image(2).Visible = False
    payment_option_image(3).Visible = False
    payment_option_image(4).Visible = True
    
    main_manu_note_label.Caption = "Note : Click here to add new electricity charge rate details."
End Sub

' add new service charge
Private Sub payment_option_new_service_Click()
    'load latest service fee details
    counter_1 = service_count_function
    Payment_Form.Service_fee_TextBox.Text = ""
    If counter_1 > 0 Then
        Open "ServiceFee.txt" For Random As #1 Len = 24
            Get #1, counter_1, service_class_1
            Payment_Form.service_add_fr_label(9).Caption = service_class_1.fee(0)
            Payment_Form.service_add_fr_label(10).Caption = service_class_1.fee(1)
            Payment_Form.service_add_fr_label(11).Caption = service_class_1.fee(2)
            Payment_Form.service_add_fr_label(12).Caption = service_class_1.fee(3)
        Close #1
    Else
    
    End If
        
    Payment_Form.fee_combo(0).Clear
    For count_1 = Format(Now, "yyyy") To Format(Now, "yyyy") + 1 Step 1
        Payment_Form.fee_combo(0).AddItem count_1
    Next count_1
    
    Payment_Form.fee_combo(0).Text = "Year"
    
    'form and frame visibility
    Payment_Form.WindowState = Main_Form.WindowState
    Payment_Form.Visible = True
    
    Payment_Form.service_add_fr.Visible = True
    Payment_Form.service_det_frame.Visible = False
    Payment_Form.electricity_fee_frame.Visible = False
    Payment_Form.payment_frame.Visible = False
    Payment_Form.payment_detail_frame.Visible = False
End Sub

Private Sub payment_option_new_service_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    payment_option_image(0).Visible = False
    payment_option_image(1).Visible = False
    payment_option_image(2).Visible = False
    payment_option_image(3).Visible = True
    payment_option_image(4).Visible = False
    
    main_manu_note_label.Caption = "Note : Click here to add new service charge details."
End Sub

Private Sub payment_option_payment_Click()
    'load room available numbers in a combo box
    available_room_count = 0

    count_2 = room_detail_count_function
    Payment_Form.payment_room_combo.Clear
    
    If count_2 = 0 Then
        MsgBox_Response = MsgBox("          It seems like you have not added any room details.", vbInformation + vbOKOnly, "Rental Record")
    Else
        Open "RoomDetail.txt" For Random As #1 Len = 112
            For count_1 = 1 To count_2 Step 1
                Get #1, count_1, existing_room
                If existing_room.room_occupied = True Then
                    Payment_Form.payment_room_combo.AddItem existing_room.room_number
                    available_room_count = available_room_count + 1
                End If
            Next count_1
        Close #1
        Payment_Form.payment_room_combo.Text = "Room number"
    End If

    Payment_Form.payment_next_command.Caption = "Check"
    Payment_Form.payment_unit_TextBox.Visible = False
    Payment_Form.payment_fr_label(3).Visible = False
    Payment_Form.payment_fr_label(30).Visible = False
    
    'form and frame visibility
    Payment_Form.WindowState = Main_Form.WindowState
    Payment_Form.Visible = True
    
    Payment_Form.payment_frame.Visible = True
    Payment_Form.service_det_frame.Visible = False
    Payment_Form.electricity_fee_frame.Visible = False
    Payment_Form.service_add_fr.Visible = False
    Payment_Form.payment_detail_frame.Visible = False
    Main_Form.Visible = False
End Sub

Private Sub payment_option_report_Click()
    'displaying payment details
    Call fill_payment_details
    
    Payment_Form.payment_filter_combo(0).Text = "Room No"
    Payment_Form.payment_filter_combo(1).Text = "Year"
    Payment_Form.payment_filter_combo(2).Text = "Month"
    
    Payment_Form.payment_filter_check(0).Value = Unchecked
    Payment_Form.payment_filter_check(1).Value = Unchecked
    Payment_Form.payment_filter_check(2).Value = Unchecked
    Payment_Form.payment_filter_check(3).Value = Unchecked
    
    Payment_Form.payment_filter_option(0).Value = True
    Payment_Form.payment_filter_command.Caption = "Filter : Off"
    Payment_Form.payment_filter_combo(0).Clear
    
    count_2 = room_detail_count_function
    Open "RoomDetail.txt" For Random As #2 Len = 112
        For count_1 = 1 To count_2
            Get #2, count_1, temp_room
            Payment_Form.payment_filter_combo(0).AddItem temp_room.room_number
        Next count_1
    Close #2
    Payment_Form.payment_filter_combo(0).Text = "Room No"
    
    Payment_Form.payment_filter_combo(1).Clear
    For i = 2022 To (return_year + 1)
        Payment_Form.payment_filter_combo(1).AddItem i
    Next i
    Payment_Form.payment_filter_combo(1).Text = "Year"
    
    'form and frame visibility
    Payment_Form.WindowState = Main_Form.WindowState
    Payment_Form.Visible = True
    
    Payment_Form.payment_detail_frame.Visible = True
    Main_Form.Visible = False
    
End Sub

'Mainmenu frame >> payment options
Private Sub payment_option_report_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    payment_option_image(0).Visible = True
    payment_option_image(1).Visible = False
    payment_option_image(2).Visible = False
    payment_option_image(3).Visible = False
    payment_option_image(4).Visible = False
    
    main_manu_note_label.Caption = "Note : Click here to show payment reports."
End Sub

Private Sub payment_option_payment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    payment_option_image(0).Visible = False
    payment_option_image(1).Visible = True
    payment_option_image(2).Visible = False
    payment_option_image(3).Visible = False
    payment_option_image(4).Visible = False
    
    main_manu_note_label.Caption = "Note : Click here to do new payment."
End Sub

Private Sub payment_option_service_Click()
    'load latest service fee details
    counter_2 = service_count_function
    counter_3 = 0
    Payment_Form.Service_fee_TextBox.Text = ""
    
    If counter_2 > 0 Then
        Open "ServiceFee.txt" For Random As #1 Len = 24
            For counter_1 = counter_2 To 1 Step -1
                Get #1, counter_1, service_class_1
                
                display_service_class.service_year(0) = service_class_1.service_year(0)
                
                Call return_month_in_string(display_service_class.service_month(0), service_class_1.service_month(0))
                
                If service_class_1.service_year(1) = 0 Then
                    display_service_class.service_year(1) = ""
                    display_service_class.service_month(1) = "Latest"
                Else
                    display_service_class.service_year(1) = CStr(service_class_1.service_year(1))
                    Call return_month_in_string(display_service_class.service_month(1), service_class_1.service_month(1))
                End If
                
                For i = 0 To 3
                    display_service_class.fee(i) = CStr(service_class_1.fee(i))
                Next i
                
                counter_3 = counter_3 + 1
                display_service_class.serial = CStr(counter_3)
                
                Payment_Form.Service_fee_TextBox.Text = Payment_Form.Service_fee_TextBox.Text + display_service_class.serial
                Payment_Form.Service_fee_TextBox.Text = Payment_Form.Service_fee_TextBox.Text + display_service_class.service_month(0) + display_service_class.service_year(0)
                Payment_Form.Service_fee_TextBox.Text = Payment_Form.Service_fee_TextBox.Text + display_service_class.service_month(1) + display_service_class.service_year(1)
                Payment_Form.Service_fee_TextBox.Text = Payment_Form.Service_fee_TextBox.Text + display_service_class.fee(0) + display_service_class.fee(1) + display_service_class.fee(2) + display_service_class.fee(3) + vbNewLine
            Next counter_1
        Close #1
    Else
        Payment_Form.Service_fee_TextBox.Text = "No service fee detail has been added till now!"
    End If
    
    'load electricity charge rate from a file
    counter_3 = 0
    counter_2 = electricity_fee_count_function
    If counter_2 > 0 Then
        Open "ElectricityFee.txt" For Random As #2 Len = 78
            Payment_Form.Electricity_fee_TextBox.Text = ""
            For counter_1 = counter_2 To 1 Step -1
                Get #2, counter_1, electricity_1
                counter_3 = counter_3 + 1
                display_electricity_class.serial = CStr(counter_3)
                
                For i = 0 To 5
                    display_electricity_class.range_min(i) = electricity_1.range_min(i)
                    display_electricity_class.range_max(i) = electricity_1.range_max(i)
                    display_electricity_class.monthly_min(i) = electricity_1.monthly_min(i)
                    display_electricity_class.per_unit(i) = electricity_1.per_unit(i)
                Next i
                
                display_electricity_class.electricity_year(0) = electricity_1.electricity_year(0)
                'display_electricity_class.electricity_month(0) = electricity_1.electricity_month(0)
                Call return_month_in_string(display_electricity_class.electricity_month(0), electricity_1.electricity_month(0))
                
                If electricity_1.electricity_year(1) = 0 Then
                    display_electricity_class.electricity_year(1) = "-"
                    display_electricity_class.electricity_month(1) = "Latest"
                Else
                    display_electricity_class.electricity_year(1) = electricity_1.electricity_year(1)
                    Call return_month_in_string(display_electricity_class.electricity_month(1), electricity_1.electricity_month(1))
                End If
                
                'displaying task
                For i = 0 To 5
                    If i > 0 Then
                        display_electricity_class.serial = ""
                        display_electricity_class.electricity_year(0) = ""
                        display_electricity_class.electricity_year(1) = ""
                        display_electricity_class.electricity_month(0) = ""
                        display_electricity_class.electricity_month(1) = ""
                    End If
                    Payment_Form.Electricity_fee_TextBox.Text = Payment_Form.Electricity_fee_TextBox.Text + display_electricity_class.serial
                    Payment_Form.Electricity_fee_TextBox.Text = Payment_Form.Electricity_fee_TextBox.Text + display_electricity_class.electricity_month(0) + display_electricity_class.electricity_year(0)
                    Payment_Form.Electricity_fee_TextBox.Text = Payment_Form.Electricity_fee_TextBox.Text + display_electricity_class.electricity_month(1) + display_electricity_class.electricity_year(1)
                    Payment_Form.Electricity_fee_TextBox.Text = Payment_Form.Electricity_fee_TextBox.Text + display_electricity_class.range_min(i) + display_electricity_class.range_max(i) + display_electricity_class.monthly_min(i)
                    Payment_Form.Electricity_fee_TextBox.Text = Payment_Form.Electricity_fee_TextBox.Text + display_electricity_class.per_unit(i) + vbNewLine
                Next i
            Next counter_1
        Close #2
    Else
        Payment_Form.Electricity_fee_TextBox.Text = "No electricity fee rate has been added till now!"
    End If
    
    'form and frame visibility
    Payment_Form.WindowState = Main_Form.WindowState
    Payment_Form.Visible = True
    
    Payment_Form.service_det_frame.Visible = True
    Payment_Form.electricity_fee_frame.Visible = False
    Payment_Form.service_add_fr.Visible = False
    Payment_Form.payment_frame.Visible = False
    Payment_Form.payment_detail_frame.Visible = False
    
    Main_Form.Visible = False
End Sub

Private Sub payment_option_service_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    payment_option_image(0).Visible = False
    payment_option_image(1).Visible = False
    payment_option_image(2).Visible = True
    payment_option_image(3).Visible = False
    payment_option_image(4).Visible = False
    
    main_manu_note_label.Caption = "Note : Click here to view service charge details."
End Sub

Private Sub payment_option_service_edit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    payment_option_image(0).Visible = False
    payment_option_image(1).Visible = False
    payment_option_image(2).Visible = False
    payment_option_image(3).Visible = True
    
    main_manu_note_label.Caption = "Note : Click here to add new service harge details."
End Sub

Private Sub remove_room_frame_combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub remove_room_frame_label5_Click()
    'frame visibility
    main_menu_frame.Visible = True
    remove_room_frame.Visible = False
End Sub

Private Sub remove_room_frame_remove_command_Click()
    'check if value if the room number is chosen or not
    If remove_room_frame_combo1.Text = "" Then
        MsgBox_Response = MsgBox("                    Please select the room number from the list first.", vbInformation + vbOKOnly, "Rental Record")
    Else
        count_1 = room_detail_count_function
        If count_1 > 0 Then
            Open "temp.txt" For Output As #1 Len = 112
            Close #1
            
            Open "RoomDetail.txt" For Random As #2 Len = 112
            Open "temp.txt" For Random As #3 Len = 112
                
            count_3 = 1
            
            new_room.room_number = remove_room_frame_combo1.Text
            
            For count_2 = 1 To count_1 Step 1
                Get #2, count_2, existing_room
                If existing_room.room_number <> new_room.room_number Then
                    Put #3, count_3, existing_room
                    count_3 = count_3 + 1
                End If
            Next count_2
            
            Close #3
            Close #2
            
            Kill "RoomDetail.txt"
            Name "temp.txt" As "RoomDetail.txt"
            
            'ask user if next room is to be removed
            MsgBox_Response = MsgBox("          Room detail removed successfully. Do you want to remove another room?", vbInformation + vbYesNo, "Room Removed")
            
            If MsgBox_Response = 6 Then
                room_option_frame_remove_Click
            Else
                main_menu_frame.Visible = True
                remove_room_frame.Visible = False
            End If
        End If
    End If
End Sub



Private Sub reset_fr_label_5_Click()
    'reset values
    LogInFrame_Timer.Enabled = True
    reset_frame_timer.Enabled = False
    
    reset_fr_text_1.Text = "Enter password"
    reset_fr_text_2.Text = "Enter password for confirmation"
    reset_fr_label_check.Value = Checked
    
    'frame visibility
    LogInFrame.Visible = True
    Reset_Frame.Visible = False
End Sub

Private Sub reset_fr_label_check_Click()
    If reset_fr_label_check.Value = Checked Then
        reset_fr_text_1.PasswordChar = ""
        reset_fr_text_2.PasswordChar = ""
    Else
        If reset_fr_text_1.Text <> "Enter password" Then reset_fr_text_1.PasswordChar = "*"
        If reset_fr_text_2.Text <> "Enter password for confirmation" Then reset_fr_text_2.PasswordChar = "*"
    End If
End Sub


Private Sub reset_fr_reset_Click()
    'check for default values
    If reset_fr_text_1.Text = "Enter password" Or reset_fr_text_2.Text = "Enter password for confirmation" Then
        If reset_fr_text_1.Text = "Enter password" Then
            MsgBox_Response = MsgBox("  Please enter valid password.", vbInformation + vbOKOnly, "Rental Record")
            reset_fr_text_1.SetFocus
        Else
            MsgBox_Response = MsgBox("  Please enter valid password for confirmation.", vbInformation + vbOKOnly, "Rental Record")
            reset_fr_text_2.SetFocus
        End If
        Exit Sub
    End If
    
    'check for password confirmation
    If reset_fr_text_1.Text = reset_fr_text_2.Text Then
        Open "AdminData.txt" For Random As #1 Len = 121
            Get #1, 1, signup
        
            Dim temp_signup As MainUser
            temp_signup.password = reset_fr_text_1.Text
        
            'check if new password match with the previous password
            If signup.password = temp_signup.password Then
                MsgBox_Response = MsgBox("  Sorry, new password matched with the previous one.", vbInformation + vbOKOnly, "Rental Record")
            Else
                signup.password = temp_signup.password
                Put #1, 1, signup
                MsgBox_Response = MsgBox("  New passwors set successfully.", vbInformation + vbOKOnly, "Rental Record")
                
                'reset values of log in frame
                LogInFrame_Password_TextBox.Text = "Enter password"
                LogInFrame_Username_TextBox.Text = "Enter username"
                LogInFrame_Timer.Enabled = True
                reset_frame_timer.Enabled = False
                login_pw_hide.Value = Unchecked
                LogInFrame_Username_TextBox.PasswordChar = ""
                LogInFrame_Password_TextBox.PasswordChar = ""
                
                reset_fr_text_1.Text = "Enter password"
                reset_fr_text_2.Text = "Enter password for confirmation"
                
                reset_fr_text_1.PasswordChar = ""
                reset_fr_text_2.PasswordChar = ""
                
                'frame visibility
                LogInFrame.Visible = True
                Reset_Frame.Visible = False
            End If
        Close #1
    Else
        MsgBox_Response = MsgBox("  Make your you entered both password correctly.", vbInformation + vbOKOnly, "Rental Record")
    End If
End Sub

Private Sub reset_fr_text_1_GotFocus()
    If reset_fr_text_1.Text = "Enter password" Then reset_fr_text_1.Text = ""
End Sub


Private Sub reset_fr_text_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If reset_fr_text_1.Text <> "" Or reset_fr_text_1.Text <> "Enter password" Then reset_fr_text_2.SetFocus
    ElseIf KeyAscii = 32 Then
        KeyAscii = 0
    Else
        If reset_fr_label_check.Value = Checked Then
            reset_fr_text_1.PasswordChar = ""
        Else
            reset_fr_text_1.PasswordChar = "*"
        End If
    End If
End Sub


Private Sub reset_fr_text_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If reset_fr_text_2.Text <> "" Or reset_fr_text_2.Text <> "Enter password for confirmation" Then reset_fr_reset_Click
    ElseIf KeyAscii = 32 Then
        KeyAscii = 0
    Else
        If reset_fr_label_check.Value = Checked Then
            reset_fr_text_2.PasswordChar = ""
        Else
            reset_fr_text_2.PasswordChar = "*"
        End If
    End If
End Sub

Private Sub reset_fr_text_2_GotFocus()
    If reset_fr_text_2.Text = "Enter password for confirmation" Then reset_fr_text_2.Text = ""
End Sub

Private Sub reset_fr_text_1_LostFocus()
    If reset_fr_text_1.Text = "" Then
        reset_fr_text_1.PasswordChar = ""
        reset_fr_text_1.Text = "Enter password"
    End If
End Sub


Private Sub reset_fr_text_2_LostFocus()
    If reset_fr_text_2.Text = "" Then
        reset_fr_text_2.PasswordChar = ""
        reset_fr_text_2.Text = "Enter password for confirmation"
    End If
End Sub


Private Sub reset_frame_timer_Timer()
    If Reset_Frame_PictureTimer < 4 Then
        Reset_Frame_PictureTimer = Reset_Frame_PictureTimer + 1
    Else
        Reset_Frame_PictureTimer = 1
    End If
    
    If Reset_Frame_PictureTimer = 1 Then
        reset_fr_image_1.Picture = LoadPicture("Assests/ApartmentPicture-1.jpg")
    ElseIf Reset_Frame_PictureTimer = 2 Then
        reset_fr_image_1.Picture = LoadPicture("Assests/ApartmentPicture-2.jpg")
    ElseIf Reset_Frame_PictureTimer = 3 Then
        reset_fr_image_1.Picture = LoadPicture("Assests/ApartmentPicture-3.jpg")
    Else
        reset_fr_image_1.Picture = LoadPicture("Assests/ApartmentPicture-4.jpg")
    End If
End Sub


Private Sub room_filter_command_Click()
    If room_filter_option(0).Value = Unchecked Then
        room_filter_status.Caption = "Filter Status : Off"
        main_menu_room_option_Click
    ElseIf room_filter_option(0).Value = Checked Then   'room status
        room_filter_status.Caption = "Filter Status : On"
    
        sn = 1
        count_2 = room_detail_count_function
        
        If count_2 = 0 Then
            room_frame_TextBox.Text = "No detail has been added yet!"
        Else
            room_frame_TextBox.Text = ""
            Open "roomdetail.txt" For Random As #1 Len = 112
                For count_1 = 1 To count_2 Step 1
                    Get #1, count_1, existing_room
                    If existing_room.room_occupied = True And room_filter_room_status(0).Value = True Then
                        room_display.sn = CStr(sn)
                        room_display.room_num = CStr(existing_room.room_number)
                        room_display.bhk = CStr(existing_room.bhk)
                        room_display.rent = CStr(existing_room.rent_amount)
                        room_display.occupied = CStr(existing_room.room_occupied)
                        room_display.tenant_fname = existing_room.tenant_fname
                        room_display.tenant_contact = existing_room.tenant_contact
                            
                        If existing_room.tenant_mname <> "Null      " Then 'has middle name
                            room_display.tenant_mname = existing_room.tenant_mname
                            room_display.tenant_lname = existing_room.tenant_lname
                        Else
                            room_display.tenant_mname = existing_room.tenant_lname
                            room_display.tenant_lname = ""
                        End If
                        
                        'for display purpose only
                        room_frame_TextBox.Text = room_frame_TextBox.Text + room_display.sn + room_display.room_num + room_display.bhk + room_display.rent + room_display.occupied
                        room_frame_TextBox.Text = room_frame_TextBox.Text + room_display.tenant_fname + room_display.tenant_mname + room_display.tenant_lname
                        room_frame_TextBox.Text = room_frame_TextBox.Text + room_display.tenant_contact + existing_room.service_provided + vbNewLine
                        sn = sn + 1
                    End If
                    
                    'unoccupied room details
                    If existing_room.room_occupied = False And room_filter_room_status(0).Value = False Then
                        room_display.sn = CStr(sn)
                        room_display.room_num = CStr(existing_room.room_number)
                        room_display.bhk = CStr(existing_room.bhk)
                        room_display.rent = CStr(existing_room.rent_amount)
                        room_display.occupied = CStr(existing_room.room_occupied)
                        room_display.tenant_fname = "-"
                        room_display.tenant_mname = "-"
                        room_display.tenant_lname = "-"
                        room_display.tenant_contact = "-"
                            
                        'for display purpose only
                        room_frame_TextBox.Text = room_frame_TextBox.Text + room_display.sn + room_display.room_num + room_display.bhk + room_display.rent + room_display.occupied
                        room_frame_TextBox.Text = room_frame_TextBox.Text + room_display.tenant_fname + room_display.tenant_lname + room_display.tenant_mname + room_display.tenant_contact
                        room_frame_TextBox.Text = room_frame_TextBox.Text + existing_room.service_provided + vbNewLine
                        sn = sn + 1
                    End If
                Next count_1
            Close #1
        End If
    End If
End Sub

Private Sub room_frame_main_label_Click()
    room_filter_option(0).Value = Unchecked
    room_filter_status.Caption = "Filter status : Off"
    
    room_frame.Visible = False
    main_menu_frame.Visible = True
End Sub

Private Sub room_option_frame_add_Click()
    'frame visibility
    LogInFrame.Visible = False
    SignUpFrame.Visible = False
    ForgotPasswordFrame.Visible = False
    main_menu_frame.Visible = False
    room_frame.Visible = False
    add_room_frame.Visible = True
End Sub


Private Sub room_option_frame_edit_Click()
    'edit frame -> loading room number in combo
    edit_room_1_combo_1.Clear
    count_1 = room_detail_count_function
    Open "RoomDetail.txt" For Random As #1 Len = 112
        For count_2 = 1 To count_1 Step 1
            Get #1, count_2, existing_room
            edit_room_1_combo_1.AddItem existing_room.room_number
        Next count_2
        edit_room_1_combo_1.Text = "Room number"
    Close #1
    
    'frame visibility
    edit_room_frame.Visible = True
    edit_room_frame_1.Visible = True
    main_menu_frame.Visible = False
    room_frame.Visible = False
    edit_room_frame_2.Visible = False
    
    edit_room_frame_save.Visible = False
    edit_room_fr_image_1.Visible = False
End Sub

Private Sub room_option_frame_remove_Click()
    remove_room_frame_combo1.Clear
    
    'fetch room number from file and add it into the combo box
    count_1 = room_detail_count_function
    Open "RoomDetail.txt" For Random As #1 Len = 112
        For count_2 = 1 To count_1 Step 1
            Get #1, count_2, existing_room
            If existing_room.room_occupied = False Then remove_room_frame_combo1.AddItem existing_room.room_number
        Next count_2
        remove_room_frame_combo1.Text = "Room number"
    Close #1
    
    'frame visibility
    remove_room_frame.Visible = True
    room_frame.Visible = False
    main_menu_frame.Visible = False
End Sub

Private Sub room_option_frame_view_Click()
    main_menu_frame.Visible = False
    main_menu_room_option_Click
End Sub

'''''sub menu frame
Private Sub room_option_frame_view_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    room_option_image(0).Visible = True
    room_option_image(1).Visible = False
    room_option_image(2).Visible = False
    room_option_image(3).Visible = False
    
    main_manu_note_label.Caption = "Note : Click here to view room details."
End Sub

Private Sub room_option_frame_add_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    room_option_image(0).Visible = False
    room_option_image(1).Visible = True
    room_option_image(2).Visible = False
    room_option_image(3).Visible = False
    
    main_manu_note_label.Caption = "Note : Click here to add new room details."
End Sub

Private Sub room_option_frame_remove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    room_option_image(0).Visible = False
    room_option_image(1).Visible = False
    room_option_image(2).Visible = True
    room_option_image(3).Visible = False
    
    main_manu_note_label.Caption = "Note : Click here to remove room details."
End Sub

Private Sub room_option_frame_edit_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    room_option_image(0).Visible = False
    room_option_image(1).Visible = False
    room_option_image(2).Visible = False
    room_option_image(3).Visible = True
    
    main_manu_note_label.Caption = "Note : Click here to edit room details."
End Sub

Private Sub signup_pw_hide_Click()
    If signup_pw_hide.Value = Checked Then
        SignUpFrame_Password_TextBox.PasswordChar = ""
        SignUpFrame_PasswordConfirmation_TextBox.PasswordChar = ""
    Else
        SignUpFrame_Password_TextBox.PasswordChar = "*"
        
        If SignUpFrame_PasswordConfirmation_TextBox.Text <> "Enter password for confirmation" Then
            SignUpFrame_PasswordConfirmation_TextBox.PasswordChar = "*"
        End If
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''Sign Up frame'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SignUpFrame_SignUp_Command_Click()
    'check for default values in textbox
    If SignUpFrame_Username_TextBox.Text = "Enter username" Or SignUpFrame_Password_TextBox.Text = "Enter password" Or SignUpFrame_PasswordConfirmation_TextBox.Text = "Enter password for confirmation" Or SignUpFrame_Contact_TextBox.Text = "Enter contact number" Or SignUpFrame_Pin_TextBox.Text = "Enter pin number" Or SignUpFrame_Security_TextBox.Text = "Enter your favourite thing" Then
        If SignUpFrame_Username_TextBox.Text = "Enter username" Then
            MsgBox_Response = MsgBox("                    Please enter the valid username.", vbInformation + vbOKOnly, "Rental Record")
            SignUpFrame_Username_TextBox.SetFocus
            Exit Sub
        ElseIf SignUpFrame_Password_TextBox.Text = "Enter password" Then
            MsgBox_Response = MsgBox("                    Please enter the valid password.", vbInformation + vbOKOnly, "Rental Record")
            SignUpFrame_Password_TextBox.SetFocus
            Exit Sub
        ElseIf SignUpFrame_PasswordConfirmation_TextBox.Text = "Enter password for confirmation" Then
            MsgBox_Response = MsgBox("                    Please enter the valid password for confirmation.", vbInformation + vbOKOnly, "Rental Record")
            SignUpFrame_PasswordConfirmation_TextBox.SetFocus
            Exit Sub
        ElseIf SignUpFrame_Contact_TextBox.Text = "Enter contact number" Then
            MsgBox_Response = MsgBox("                    Please enter the valid contact number.", vbInformation + vbOKOnly, "Rental Record")
            SignUpFrame_Contact_TextBox.SetFocus
            Exit Sub
        ElseIf SignUpFrame_Pin_TextBox.Text = "Enter pin number" Then
            MsgBox_Response = MsgBox("                    Please enter the valid pin number.", vbInformation + vbOKOnly, "Rental Record")
            SignUpFrame_Pin_TextBox.SetFocus
            Exit Sub
        Else
            MsgBox_Response = MsgBox("                    Please enter the valid security answer.", vbInformation + vbOKOnly, "Rental Record")
            SignUpFrame_Security_TextBox.SetFocus
            Exit Sub
        End If
    End If
    
    'check values of combo bix allocated for date of birth
    If SignUpFrame_dob_Combo(0).Text = "Year" Or SignUpFrame_dob_Combo(1).Text = "Month" Or SignUpFrame_dob_Combo(2).Text = "Date" Then
        If SignUpFrame_dob_Combo(0).Text = "Year" Then
            MsgBox_Response = MsgBox("                    Please set the your birth year.", vbInformation + vbOKOnly, "Rental Record")
            SignUpFrame_dob_Combo(0).SetFocus
            Exit Sub
        ElseIf SignUpFrame_dob_Combo(1).Text = "Month" Then
            MsgBox_Response = MsgBox("                    Please set the your of birth month.", vbInformation + vbOKOnly, "Rental Record")
            SignUpFrame_dob_Combo(1).SetFocus
            Exit Sub
        Else
            MsgBox_Response = MsgBox("                    Please set the your birth date.", vbInformation + vbOKOnly, "Rental Record")
            SignUpFrame_dob_Combo(2).SetFocus
            Exit Sub
        End If
    End If
    
    'save data in a file
    signup.username = SignUpFrame_Username_TextBox.Text
    signup.password = SignUpFrame_Password_TextBox.Text
    signup.contact_number = SignUpFrame_Contact_TextBox.Text
    signup.pin_number = SignUpFrame_Pin_TextBox.Text
    signup.security_question = SignUpFrame_Security_TextBox.Text
    signup.dob_year = SignUpFrame_dob_Combo(0).Text
    signup.dob_month = SignUpFrame_dob_Combo(1).Text
    signup.dob_date = SignUpFrame_dob_Combo(2).Text
    
    'check for password confirmation
    If SignUpFrame_Password_TextBox.Text = SignUpFrame_PasswordConfirmation_TextBox Then
        'write admin details in a file
        Open "AdminData.txt" For Random As #1 Len = 121
            Put #1, 1, signup
        Close #1
        
        MsgBox_Response = MsgBox("                    Congratulations! you have successfully signed up.                     Click on OK button to continue to the application.", vbInformation + vbOKOnly, "Rental Record")
        
        signup_pw_hide.Value = Checked
        SignUpFrame_Password_TextBox.PasswordChar = ""
        SignUpFrame_PasswordConfirmation_TextBox.PasswordChar = ""
        
        LogInFrame_Timer.Enabled = True
        SignUpFrame_Timer.Enabled = False
        ForgotPasswordFrame_Timer.Enabled = False
        
        
        LogInFrame.Visible = True
        SignUpFrame.Visible = False
        ForgotPasswordFrame.Visible = False
    Else
        MsgBox_Response = MsgBox("                    Make sure you entered the passwords same.", vbInformation + vbOKOnly, "Rental Record")
        If MsgBox_Response = 1 Then SignUpFrame_PasswordConfirmation_TextBox.SetFocus
    End If
End Sub


'username box
Private Sub SignUpFrame_Username_TextBox_GotFocus()
    If SignUpFrame_Username_TextBox.Text = "Enter username" Then SignUpFrame_Username_TextBox.Text = ""
End Sub

Private Sub SignUpFrame_Username_TextBox_LostFocus()
    If SignUpFrame_Username_TextBox.Text = "" Then SignUpFrame_Username_TextBox.Text = "Enter username"
End Sub

Private Sub SignUpFrame_Username_TextBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And SignUpFrame_Username_TextBox.Text <> "" And SignUpFrame_Username_TextBox <> "Enter username" Then
        KeyAscii = 0
        SignUpFrame_Password_TextBox.SetFocus
    End If
End Sub


'password box
Private Sub SignUpFrame_Password_TextBox_GotFocus()
    If SignUpFrame_Password_TextBox.Text = "Enter password" Then
        signup_pw_hide.Value = Unchecked
        SignUpFrame_Password_TextBox.Text = ""
        SignUpFrame_Password_TextBox.PasswordChar = "*"
        
        If SignUpFrame_PasswordConfirmation_TextBox.Text = "Enter password for confirmation" Then
            SignUpFrame_PasswordConfirmation_TextBox.PasswordChar = ""
        End If
    End If
End Sub

Private Sub SignUpFrame_Password_TextBox_LostFocus()
    If SignUpFrame_Password_TextBox.Text = "" Then
        SignUpFrame_Password_TextBox.Text = "Enter password"
        SignUpFrame_PasswordConfirmation_TextBox.Text = "Enter password for confirmation"
        signup_pw_hide.Value = Checked
    End If
End Sub

Private Sub SignUpFrame_Password_TextBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And SignUpFrame_Password_TextBox.Text <> "" And SignUpFrame_Password_TextBox.Text <> "Enter password" Then
        KeyAscii = 0
        SignUpFrame_PasswordConfirmation_TextBox.SetFocus
    End If
End Sub


'password confirmation box
Private Sub SignUpFrame_PasswordConfirmation_TextBox_GotFocus()
    If SignUpFrame_PasswordConfirmation_TextBox.Text = "Enter password for confirmation" Then
        SignUpFrame_PasswordConfirmation_TextBox.PasswordChar = "*"
        SignUpFrame_PasswordConfirmation_TextBox.Text = ""
    End If
End Sub

Private Sub SignUpFrame_PasswordConfirmation_TextBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And SignUpFrame_PasswordConfirmation_TextBox.Text <> "" And SignUpFrame_PasswordConfirmation_TextBox.Text <> "Enter password for confirmation" Then
        KeyAscii = 0
        SignUpFrame_Contact_TextBox.SetFocus
    End If
End Sub

Private Sub SignUpFrame_PasswordConfirmation_TextBox_LostFocus()
    If SignUpFrame_PasswordConfirmation_TextBox.Text = "" Then
        SignUpFrame_PasswordConfirmation_TextBox.Text = "Enter password for confirmation"
        SignUpFrame_PasswordConfirmation_TextBox.PasswordChar = ""
    End If
End Sub


'contact number box
Private Sub SignUpFrame_Contact_TextBox_GotFocus()
    If SignUpFrame_Contact_TextBox.Text = "Enter contact number" Then SignUpFrame_Contact_TextBox.Text = ""
End Sub

Private Sub SignUpFrame_Contact_TextBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    
    If KeyAscii = 13 And SignUpFrame_Contact_TextBox.Text <> "" And SignUpFrame_Contact_TextBox.Text <> "Enter contact number" Then
        KeyAscii = 0
        SignUpFrame_Pin_TextBox.SetFocus
    End If
    
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub SignUpFrame_Contact_TextBox_LostFocus()
    If SignUpFrame_Contact_TextBox.Text = "" Then SignUpFrame_Contact_TextBox.Text = "Enter contact number"
End Sub

'pin number of Sign Up Frame
Private Sub SignUpFrame_Pin_TextBox_GotFocus()
    If SignUpFrame_Pin_TextBox.Text = "Enter pin number" Then SignUpFrame_Pin_TextBox.Text = ""
End Sub

Private Sub SignUpFrame_Pin_TextBox_LostFocus()
    If SignUpFrame_Pin_TextBox.Text = "" Then SignUpFrame_Pin_TextBox.Text = "Enter pin number"
End Sub

'security question textbox
Private Sub ForgotPasswordFrame_security_TextBox_GotFocus()
    If ForgotPasswordFrame_Security_TextBox.Text = "Enter your favourite thing" Then ForgotPasswordFrame_Security_TextBox.Text = ""
End Sub

Private Sub ForgotPasswordFrame_security_TextBox_LostFocus()
    If ForgotPasswordFrame_Security_TextBox.Text = "" Then ForgotPasswordFrame_Security_TextBox.Text = "Enter your favourite thing"
End Sub

Private Sub ForgotPasswordFrame_Security_TextBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And ForgotPasswordFrame_Security_TextBox.Text <> "" Then
        KeyAscii = 0
        ForgotPasswordFrame_dob_Combo(0).SetFocus
    End If
End Sub

Private Sub SignUpFrame_Pin_TextBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    
    If KeyAscii <> 13 Then
        If KeyAscii > 46 And KeyAscii < 63 Then
            KeyAscii = KeyAscii
        Else
            KeyAscii = 0
        End If
    ElseIf KeyAscii = 13 And SignUpFrame_Pin_TextBox.Text <> "" And SignUpFrame_Pin_TextBox.Text <> "Enter contact number" Then
        KeyAscii = 0
        SignUpFrame_dob_Combo(0).SetFocus
    End If
End Sub


'security question text box
Private Sub SignUpFrame_Security_TextBox_GotFocus()
    If SignUpFrame_Security_TextBox.Text = "Enter your favourite thing" Then SignUpFrame_Security_TextBox.Text = ""
End Sub

Private Sub SignUpFrame_Security_TextBox_LostFocus()
    If SignUpFrame_Security_TextBox.Text = "" Then SignUpFrame_Security_TextBox.Text = "Enter your favourite thing"
End Sub

Private Sub SignUpFrame_Security_TextBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And SignUpFrame_Security_TextBox.Text <> "Enter your favourite thing" Then
        KeyAscii = 0
        SignUpFrame_SignUp_Command_Click
    End If
End Sub

'combo box
Private Sub SignUpFrame_dob_Combo_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        If KeyAscii = 8 Then
            KeyAscii = KeyAscii
            Exit Sub
        End If
        
        If KeyAscii > 47 And KeyAscii < 58 Then
            KeyAscii = KeyAscii
            Exit Sub
        ElseIf KeyAscii <> 13 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    If KeyAscii = 13 Then
        If Index = 0 And SignUpFrame_dob_Combo(0).Text <> "Year" Then SignUpFrame_dob_Combo(1).SetFocus
        If Index = 1 And SignUpFrame_dob_Combo(1).Text <> "Month" Then SignUpFrame_dob_Combo(2).SetFocus
        If Index = 2 And SignUpFrame_dob_Combo(2).Text <> "Date" Then SignUpFrame_Security_TextBox.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub SignUpFrame_Timer_Timer()
    'change counting value
    If SignUpFrame_PictureTimer < 4 Then
        SignUpFrame_PictureTimer = SignUpFrame_PictureTimer + 1
    Else
        SignUpFrame_PictureTimer = 1
    End If
    
    'changing picture
    If SignUpFrame_PictureTimer = 1 Then
        SignUpFrameImage1.Picture = LoadPicture("Assests/ApartmentPicture-1.jpg")
    ElseIf SignUpFrame_PictureTimer = 2 Then
        SignUpFrameImage1.Picture = LoadPicture("Assests/ApartmentPicture-2.jpg")
    ElseIf SignUpFrame_PictureTimer = 3 Then
        SignUpFrameImage1.Picture = LoadPicture("Assests/ApartmentPicture-3.jpg")
    Else
        SignUpFrameImage1.Picture = LoadPicture("Assests/ApartmentPicture-4.jpg")
    End If
End Sub

Private Sub tenant_option_frame_add_Click()
    'load room available numbers in a combo box
    available_room_count = 0

    count_2 = room_detail_count_function
    
    Tenant_Form.add_tenant_fr_combo_2.Clear
    If count_2 = 0 Then
        MsgBox_Response = MsgBox("          It seems like you have not added any room details.", vbInformation + vbOKOnly, "Rental Record")
    Else
        Open "RoomDetail.txt" For Random As #1 Len = 112
            For count_1 = 1 To count_2 Step 1
                Get #1, count_1, existing_room
                If existing_room.room_occupied = False Then
                    Tenant_Form.add_tenant_fr_combo_2.AddItem existing_room.room_number
                    vailable_room_count = available_room_count + 1
                End If
            Next count_1
            Tenant_Form.add_tenant_fr_combo_2.Text = "Room number"
            
            Tenant_Form.add_tenant_fr_combo_3(0).Clear
            
            For count_3 = Format(Now, "yyyy") To Format(Now, "yyyy") + 1 Step 1
                Tenant_Form.add_tenant_fr_combo_3(0).AddItem count_3
            Next count_3
            Tenant_Form.add_tenant_fr_combo_3(0).Text = "Year"
        Close #1
    End If
    
    Tenant_Form.WindowState = Main_Form.WindowState
    Tenant_Form.Visible = True
    Main_Form.Visible = False
    tenant_option_frame.Visible = False
    
    Tenant_Form.tenant_frame.Visible = False
    Tenant_Form.add_tenant_frame.Visible = True
    Tenant_Form.remove_tenant_frame.Visible = False
    Tenant_Form.edit_tenant_frame.Visible = False
End Sub

Private Sub tenant_option_frame_edit_Click()
    'load occupied room numbers in a combo box of edit tenant frame
    Tenant_Form.edit_tenant_fr_1_Combo1.Clear
    
    count_2 = room_detail_count_function
    If count_2 > 0 Then
        Open "RoomDetail.txt" For Random As #1 Len = 112
            For count_1 = 1 To count_2 Step 1
                Get #1, count_1, existing_room
                If existing_room.room_occupied = True Then
                    Tenant_Form.edit_tenant_fr_1_Combo1.AddItem existing_room.room_number
                Else
                    count_2 = count_2
                    Tenant_Form.edit_tenant_fr_combo_1.AddItem existing_room.room_number
                End If
            Next count_1
            Tenant_Form.edit_tenant_fr_1_Combo1.Text = "Room number"
        Close #1
        
        'frame visibility
        Tenant_Form.WindowState = Main_Form.WindowState
        Tenant_Form.Visible = True
        Tenant_Form.edit_tenant_frame.Visible = True
        Tenant_Form.tenant_frame.Visible = False
        Tenant_Form.add_tenant_frame.Visible = False
        Tenant_Form.remove_tenant_frame.Visible = False
        tenant_option_frame.Visible = False
        
        Tenant_Form.edit_tenant_fr_1.Visible = True
        Tenant_Form.edit_tenant_fr_2.Visible = False
        Tenant_Form.edit_room_frame_save.Visible = False
        Tenant_Form.edit_tenant_fr_image_1.Visible = False
        Main_Form.Visible = False
    Else
        MsgBox_Response = MsgBox("     Sorry no rooms are available currently.", vbInformation + vbOKOnly, "Rental Record")
    End If
End Sub

Private Sub tenant_option_frame_remove_Click()
    Tenant_Form.remove_tenant_fr_remove_command.Caption = "Remove"
    'load occupied room numbers in a combo box of tenant form -> remove tenant frame -> combo
    Tenant_Form.remove_tenant_fr_combo_1.Clear
    count_2 = room_detail_count_function
    
    Open "RoomDetail.txt" For Random As #1 Len = 112
        For count_1 = 1 To count_2 Step 1
            Get #1, count_1, existing_room
            If existing_room.room_occupied = True Then Tenant_Form.remove_tenant_fr_combo_1.AddItem existing_room.room_number
        Next count_1
        Tenant_Form.remove_tenant_fr_combo_1.Text = "Room number"
    Close #1
    
    'form and frame visibility
    Tenant_Form.WindowState = Main_Form.WindowState
    Tenant_Form.Visible = True
    Tenant_Form.remove_tenant_frame.Visible = True
    Main_Form.Visible = False
    
    Tenant_Form.tenant_frame.Visible = False
    Tenant_Form.add_tenant_frame.Visible = False
    Tenant_Form.edit_tenant_frame.Visible = False
End Sub

Private Sub tenant_option_frame_view_Click()
    Tenant_Form.tenant_fr_TextBox.Text = ""
    
    count_2 = tenant_detail_count_function
    
    If count_2 = 0 Then
        Tenant_Form.tenant_fr_TextBox.Text = "No detail has been added yet!"
    Else
        Dim display_tenant As tenant_class_display
        
        Tenant_Form.tenant_filter_option_status(0).Value = Checked
        
        count_3 = room_detail_count_function
        Tenant_Form.tenant_filter_option_room.Clear ' clear room numbers
        Open "RoomDetail.txt" For Random As #1 Len = 112
            For i = 1 To count_3
                Get #1, i, temp_room
                Tenant_Form.tenant_filter_option_room.AddItem temp_room.room_number
            Next i
        Close #1
        
        Open "TenantDetail.txt" For Random As #2 Len = 103
            For count_1 = 1 To count_2 Step 1
                Get #2, count_1, tenant_1
                display_tenant.serial = CStr(count_1)
                display_tenant.room_num = CStr(tenant_1.room_num)
                
                
                display_tenant.first_name = tenant_1.first_name
                display_tenant.address_district = tenant_1.address_district
                display_tenant.address_municipality = tenant_1.address_municipality
                display_tenant.address_ward = tenant_1.address_ward
                display_tenant.citizenship = tenant_1.citizenship
                display_tenant.contact_num = tenant_1.contact_num
                display_tenant.rent_year(0) = tenant_1.rent_year(0)
                display_tenant.rent_year(1) = tenant_1.rent_year(1)
                
                
                Call return_month_in_string(display_tenant.rent_month(0), tenant_1.rent_month(0))
                Call return_month_in_string(display_tenant.rent_month(1), tenant_1.rent_month(1))
                
                If tenant_1.middle_name = "Null      " Then ' has no middle name
                    display_tenant.middle_name = tenant_1.last_name
                    display_tenant.last_name = ""
                Else
                    display_tenant.middle_name = tenant_1.middle_name
                    display_tenant.last_name = tenant_1.last_name
                End If
                
                Tenant_Form.tenant_fr_TextBox.Text = Tenant_Form.tenant_fr_TextBox.Text + display_tenant.serial + display_tenant.room_num
                Tenant_Form.tenant_fr_TextBox.Text = Tenant_Form.tenant_fr_TextBox.Text + display_tenant.first_name + display_tenant.middle_name + display_tenant.last_name
                Tenant_Form.tenant_fr_TextBox.Text = Tenant_Form.tenant_fr_TextBox.Text + display_tenant.address_district + display_tenant.address_municipality + display_tenant.address_ward
                Tenant_Form.tenant_fr_TextBox.Text = Tenant_Form.tenant_fr_TextBox.Text + display_tenant.contact_num + display_tenant.citizenship
                Tenant_Form.tenant_fr_TextBox.Text = Tenant_Form.tenant_fr_TextBox.Text + display_tenant.rent_year(0) + display_tenant.rent_month(0)
                
                If tenant_1.rent_year(1) = 0 And tenant_1.rent_month(1) = 0 Then 'still living
                    Tenant_Form.tenant_fr_TextBox.Text = Tenant_Form.tenant_fr_TextBox.Text + "Still living"
                Else
                    Tenant_Form.tenant_fr_TextBox.Text = Tenant_Form.tenant_fr_TextBox.Text + display_tenant.rent_year(1) + display_tenant.rent_month(1)
                End If
                Tenant_Form.tenant_fr_TextBox.Text = Tenant_Form.tenant_fr_TextBox.Text + vbNewLine
            Next count_1
        Close #2
    End If
    
    'form and frame visibility
    Tenant_Form.WindowState = Main_Form.WindowState
    Tenant_Form.Visible = True
    Tenant_Form.tenant_frame.Visible = True
    
    Main_Form.Visible = False
    Tenant_Form.add_tenant_frame.Visible = False
    Tenant_Form.remove_tenant_frame.Visible = False
    Tenant_Form.edit_tenant_frame.Visible = False
    tenant_option_frame.Visible = False
End Sub

Private Sub tenant_option_frame_view_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tenant_option_image(0).Visible = True
    tenant_option_image(1).Visible = False
    tenant_option_image(2).Visible = False
    tenant_option_image(3).Visible = False
    
    main_manu_note_label.Caption = "Note : Click here to view tenant details."
End Sub

Private Sub tenant_option_frame_add_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tenant_option_image(0).Visible = False
    tenant_option_image(1).Visible = True
    tenant_option_image(2).Visible = False
    tenant_option_image(3).Visible = False
    
    main_manu_note_label.Caption = "Note : Click here to addnew tenant details."
End Sub

Private Sub tenant_option_frame_remove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tenant_option_image(0).Visible = False
    tenant_option_image(1).Visible = False
    tenant_option_image(2).Visible = True
    tenant_option_image(3).Visible = False
    
    main_manu_note_label.Caption = "Note : Click here to remove tenant."
End Sub

Private Sub tenant_option_frame_edit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tenant_option_image(0).Visible = False
    tenant_option_image(1).Visible = False
    tenant_option_image(2).Visible = False
    tenant_option_image(3).Visible = True
    
    main_manu_note_label.Caption = "Note : Click here to edit tenant details."
End Sub
