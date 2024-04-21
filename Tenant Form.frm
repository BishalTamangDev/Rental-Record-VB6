VERSION 5.00
Begin VB.Form Tenant_Form 
   BackColor       =   &H80000009&
   Caption         =   "Rental Record"
   ClientHeight    =   12375
   ClientLeft      =   -45
   ClientTop       =   300
   ClientWidth     =   22800
   LinkTopic       =   "Form2"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame edit_tenant_frame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Edit Tenant Detail frame"
      Height          =   12375
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   23055
      Begin VB.Frame edit_tenant_fr_1 
         BackColor       =   &H8000000E&
         Height          =   2895
         Left            =   8513
         TabIndex        =   44
         Top             =   5400
         Width           =   5775
         Begin VB.ComboBox edit_tenant_fr_1_Combo1 
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
            Left            =   2640
            TabIndex        =   46
            Text            =   "Combo1"
            Top             =   1200
            Width           =   2895
         End
         Begin VB.CommandButton edit_tenant_fr_1_command_1 
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
            TabIndex        =   45
            Top             =   2040
            Width           =   5295
         End
         Begin VB.Label edit_tenant_fr_label 
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
            Index           =   20
            Left            =   240
            TabIndex        =   47
            Top             =   1200
            Width           =   1470
         End
         Begin VB.Label edit_tenant_fr_label 
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
            Height          =   975
            Index           =   19
            Left            =   240
            TabIndex        =   48
            Top             =   240
            Width           =   5145
         End
      End
      Begin VB.Frame edit_tenant_fr_2 
         BackColor       =   &H8000000E&
         Height          =   6495
         Left            =   5513
         TabIndex        =   49
         Top             =   3480
         Width           =   11775
         Begin VB.TextBox edit_tenant_fr_textbox_1 
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
            Index           =   4
            Left            =   6720
            TabIndex        =   79
            Text            =   "Enter citizenship number"
            Top             =   5520
            Width           =   4695
         End
         Begin VB.TextBox edit_tenant_fr_textbox_1 
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
            Left            =   6720
            TabIndex        =   78
            Text            =   "Enter contact number"
            Top             =   4800
            Width           =   4695
         End
         Begin VB.ComboBox edit_tenant_fr_combo_1 
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
            Left            =   6720
            TabIndex        =   77
            Text            =   "Room Number"
            Top             =   1200
            Width           =   4695
         End
         Begin VB.ComboBox edit_tenant_fr_combo_2 
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
            ItemData        =   "Tenant Form.frx":0000
            Left            =   6720
            List            =   "Tenant Form.frx":0022
            TabIndex        =   76
            Text            =   "Ward Number"
            Top             =   4080
            Width           =   4695
         End
         Begin VB.ComboBox edit_tenant_fr_combo_2 
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
            ItemData        =   "Tenant Form.frx":0045
            Left            =   6720
            List            =   "Tenant Form.frx":08B3
            TabIndex        =   75
            Text            =   "Municipality"
            Top             =   3360
            Width           =   4695
         End
         Begin VB.ComboBox edit_tenant_fr_combo_2 
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
            ItemData        =   "Tenant Form.frx":281D
            Left            =   6720
            List            =   "Tenant Form.frx":2902
            TabIndex        =   74
            Text            =   "District"
            Top             =   2640
            Width           =   4695
         End
         Begin VB.TextBox edit_tenant_fr_textbox_1 
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
            Left            =   9840
            TabIndex        =   52
            Text            =   "Last Name"
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox edit_tenant_fr_textbox_1 
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
            Left            =   8160
            TabIndex        =   51
            Text            =   "Middle Name"
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox edit_tenant_fr_textbox_1 
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
            Left            =   6720
            TabIndex        =   50
            Text            =   "First Name"
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label edit_tenant_fr_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "xxxxxxxxxxxxxxxxxxxx"
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
            Index           =   18
            Left            =   2640
            TabIndex        =   73
            Top             =   5520
            Width           =   1800
         End
         Begin VB.Label edit_tenant_fr_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "xxxxxxxxxx"
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
            Index           =   17
            Left            =   2640
            TabIndex        =   72
            Top             =   4680
            Width           =   900
         End
         Begin VB.Label edit_tenant_fr_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "xx"
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
            Index           =   16
            Left            =   4080
            TabIndex        =   71
            Top             =   3960
            Width           =   180
         End
         Begin VB.Label edit_tenant_fr_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Ward            :"
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
            Index           =   15
            Left            =   2640
            TabIndex        =   70
            Top             =   3960
            Width           =   1245
         End
         Begin VB.Label edit_tenant_fr_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "xxxxxxxxxxxxxx"
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
            Index           =   14
            Left            =   4080
            TabIndex        =   69
            Top             =   3360
            Width           =   1260
         End
         Begin VB.Label edit_tenant_fr_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Municipality  :"
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
            Index           =   13
            Left            =   2640
            TabIndex        =   68
            Top             =   3360
            Width           =   1200
         End
         Begin VB.Label edit_tenant_fr_label 
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
            Index           =   9
            Left            =   2640
            TabIndex        =   67
            Top             =   1200
            Width           =   90
         End
         Begin VB.Label edit_tenant_fr_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Citizenship Number"
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
            Index           =   8
            Left            =   240
            TabIndex        =   66
            Top             =   5520
            Width           =   1710
         End
         Begin VB.Label edit_tenant_fr_label 
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
            TabIndex        =   61
            Top             =   1200
            Width           =   1305
         End
         Begin VB.Label edit_tenant_fr_label 
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
            Index           =   7
            Left            =   240
            TabIndex        =   60
            Top             =   4680
            Width           =   1440
         End
         Begin VB.Label edit_tenant_fr_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
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
            TabIndex        =   59
            Top             =   2640
            Width           =   720
         End
         Begin VB.Label edit_tenant_fr_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
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
            TabIndex        =   58
            Top             =   1920
            Width           =   525
         End
         Begin VB.Label edit_tenant_fr_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "xxxxx"
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
            Left            =   2640
            TabIndex        =   57
            Top             =   1920
            Width           =   450
         End
         Begin VB.Label edit_tenant_fr_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "District          :"
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
            Left            =   2640
            TabIndex        =   56
            Top             =   2640
            Width           =   1245
         End
         Begin VB.Label edit_tenant_fr_label 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "xxxxxxxxxxxxxx"
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
            Left            =   4080
            TabIndex        =   55
            Top             =   2640
            Width           =   1260
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            X1              =   2400
            X2              =   2400
            Y1              =   120
            Y2              =   6480
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000010&
            X1              =   6480
            X2              =   6480
            Y1              =   120
            Y2              =   6480
         End
         Begin VB.Label edit_tenant_fr_label 
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
            Left            =   2640
            TabIndex        =   54
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label edit_tenant_fr_label 
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
            Left            =   6720
            TabIndex        =   53
            Top             =   240
            Width           =   1185
         End
         Begin VB.Line Line9 
            BorderColor     =   &H80000010&
            X1              =   0
            X2              =   11520
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Image Image2 
            Height          =   615
            Left            =   0
            Picture         =   "Tenant Form.frx":2BBB
            Stretch         =   -1  'True
            Top             =   120
            Width           =   11775
         End
      End
      Begin VB.Label edit_tenant_fr_label 
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
         TabIndex        =   65
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label edit_tenant_fr_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Edit Tenant Detail"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   17.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   465
         Index           =   1
         Left            =   960
         TabIndex        =   64
         Top             =   2280
         Width           =   2760
      End
      Begin VB.Image Image23 
         Height          =   1050
         Left            =   5520
         Picture         =   "Tenant Form.frx":30C5
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label edit_room_frame_go_back 
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
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   1200
         TabIndex        =   63
         Top             =   10560
         Width           =   1215
      End
      Begin VB.Label edit_room_frame_save 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save"
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
         Left            =   3480
         TabIndex        =   62
         Top             =   10560
         Width           =   600
      End
      Begin VB.Image Image15 
         Appearance      =   0  'Flat
         Height          =   690
         Left            =   720
         Picture         =   "Tenant Form.frx":CDCF
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   21345
      End
      Begin VB.Image Image8 
         Height          =   615
         Left            =   720
         Picture         =   "Tenant Form.frx":D2D9
         Stretch         =   -1  'True
         Top             =   10440
         Width           =   1935
      End
      Begin VB.Image edit_tenant_fr_image_1 
         Height          =   615
         Left            =   2880
         Picture         =   "Tenant Form.frx":D7E3
         Stretch         =   -1  'True
         Top             =   10440
         Width           =   1935
      End
   End
   Begin VB.Frame add_tenant_frame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Add Tenant Frame"
      Height          =   12375
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   23055
      Begin VB.CheckBox add_tenant_fr_check 
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
         Left            =   10080
         TabIndex        =   111
         Top             =   9600
         Width           =   2415
      End
      Begin VB.CheckBox add_tenant_fr_check 
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
         Left            =   12240
         TabIndex        =   110
         Top             =   9120
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox add_tenant_fr_check 
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
         Index           =   2
         Left            =   12240
         TabIndex        =   109
         Top             =   8640
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox add_tenant_fr_check 
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
         Height          =   495
         Index           =   1
         Left            =   10080
         TabIndex        =   108
         Top             =   9120
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox add_tenant_fr_check 
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
         Index           =   0
         Left            =   10080
         TabIndex        =   107
         Top             =   8640
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.ComboBox add_tenant_fr_combo_3 
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
         ItemData        =   "Tenant Form.frx":DCEE
         Left            =   12240
         List            =   "Tenant Form.frx":DD16
         TabIndex        =   104
         Text            =   "Month"
         Top             =   8040
         Width           =   2295
      End
      Begin VB.ComboBox add_tenant_fr_combo_3 
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
         ItemData        =   "Tenant Form.frx":DD7C
         Left            =   10080
         List            =   "Tenant Form.frx":DDCE
         TabIndex        =   103
         Text            =   "Year"
         Top             =   8040
         Width           =   2055
      End
      Begin VB.TextBox add_tenant_fr_textbox 
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
         Left            =   13080
         TabIndex        =   42
         Text            =   "Last Name"
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox add_tenant_fr_textbox 
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
         Left            =   11640
         TabIndex        =   41
         Text            =   "Middle Name"
         Top             =   5160
         Width           =   1335
      End
      Begin VB.ComboBox add_tenant_fr_Combo_1 
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
         ItemData        =   "Tenant Form.frx":DE6E
         Left            =   13680
         List            =   "Tenant Form.frx":DE90
         TabIndex        =   40
         Text            =   "Ward"
         Top             =   5880
         Width           =   855
      End
      Begin VB.ComboBox add_tenant_fr_Combo_1 
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
         ItemData        =   "Tenant Form.frx":DEB3
         Left            =   11880
         List            =   "Tenant Form.frx":E721
         TabIndex        =   39
         Text            =   "Municipality"
         Top             =   5880
         Width           =   1710
      End
      Begin VB.ComboBox add_tenant_fr_Combo_1 
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
         ItemData        =   "Tenant Form.frx":1068B
         Left            =   10080
         List            =   "Tenant Form.frx":10770
         TabIndex        =   38
         Text            =   "District"
         Top             =   5880
         Width           =   1710
      End
      Begin VB.TextBox add_tenant_fr_textbox 
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
         Index           =   4
         Left            =   10080
         TabIndex        =   37
         Text            =   "Enter citizenship number"
         Top             =   7320
         Width           =   4455
      End
      Begin VB.ComboBox add_tenant_fr_combo_2 
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
         Left            =   10080
         Sorted          =   -1  'True
         TabIndex        =   35
         Text            =   "Room number"
         Top             =   4440
         Width           =   4455
      End
      Begin VB.TextBox add_tenant_fr_textbox 
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
         Left            =   10080
         TabIndex        =   27
         Text            =   "First Name"
         Top             =   5160
         Width           =   1455
      End
      Begin VB.CommandButton add_tenant_fr_add_command 
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
         Left            =   7920
         TabIndex        =   26
         Top             =   10200
         Width           =   6855
      End
      Begin VB.TextBox add_tenant_fr_textbox 
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
         Left            =   10080
         TabIndex        =   25
         Text            =   "Enter contact number"
         Top             =   6600
         Width           =   4455
      End
      Begin VB.Label add_tenant_fr_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Service"
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
         Index           =   8
         Left            =   8160
         TabIndex        =   106
         Top             =   8760
         Width           =   615
      End
      Begin VB.Label add_tenant_fr_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   8160
         TabIndex        =   105
         Top             =   8040
         Width           =   405
      End
      Begin VB.Label add_tenant_fr_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Citizenship Number"
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
         Left            =   8160
         TabIndex        =   36
         Top             =   7320
         Width           =   1710
      End
      Begin VB.Label add_tenant_fr_goback 
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
         TabIndex        =   34
         Top             =   11160
         Width           =   885
      End
      Begin VB.Label add_tenant_fr_label 
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
         TabIndex        =   33
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label add_tenant_fr_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Add Tenant Details"
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
         Left            =   9900
         TabIndex        =   32
         Top             =   3000
         Width           =   2730
      End
      Begin VB.Label add_tenant_fr_label 
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
         Left            =   8160
         TabIndex        =   31
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label add_tenant_fr_label 
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
         Index           =   5
         Left            =   8160
         TabIndex        =   30
         Top             =   6600
         Width           =   1440
      End
      Begin VB.Label add_tenant_fr_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   8160
         TabIndex        =   29
         Top             =   5880
         Width           =   720
      End
      Begin VB.Label add_tenant_fr_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   8160
         TabIndex        =   28
         Top             =   5160
         Width           =   525
      End
      Begin VB.Image Image22 
         Height          =   1050
         Left            =   5520
         Picture         =   "Tenant Form.frx":10A29
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Image Image14 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   7680
         Picture         =   "Tenant Form.frx":1A733
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   7335
      End
      Begin VB.Image Image3 
         Height          =   615
         Left            =   7920
         Picture         =   "Tenant Form.frx":1AC3D
         Stretch         =   -1  'True
         Top             =   11040
         Width           =   6855
      End
      Begin VB.Image Image6 
         Height          =   12015
         Left            =   5520
         Picture         =   "Tenant Form.frx":1B148
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   11655
      End
   End
   Begin VB.Frame remove_tenant_frame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "s"
      ForeColor       =   &H80000008&
      Height          =   12375
      Left            =   0
      TabIndex        =   80
      Top             =   0
      Width           =   23055
      Begin VB.TextBox remove_tenant_fr_TextBox 
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
         Left            =   10200
         TabIndex        =   90
         Text            =   "First Name"
         Top             =   6240
         Width           =   1335
      End
      Begin VB.TextBox remove_tenant_fr_TextBox 
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
         Left            =   10200
         TabIndex        =   89
         Text            =   "Enter contact number"
         Top             =   7080
         Width           =   4335
      End
      Begin VB.TextBox remove_tenant_fr_TextBox 
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
         Left            =   13320
         TabIndex        =   88
         Text            =   "Last Name"
         Top             =   6240
         Width           =   1215
      End
      Begin VB.TextBox remove_tenant_fr_TextBox 
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
         Left            =   11640
         TabIndex        =   87
         Text            =   "Middle Name"
         Top             =   6240
         Width           =   1575
      End
      Begin VB.ComboBox remove_tenant_fr_combo_1 
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
         ItemData        =   "Tenant Form.frx":21894
         Left            =   10200
         List            =   "Tenant Form.frx":218A7
         TabIndex        =   82
         Text            =   "Room number"
         Top             =   5400
         Width           =   4335
      End
      Begin VB.CommandButton remove_tenant_fr_remove_command 
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
         Left            =   8400
         TabIndex        =   81
         Top             =   7920
         Width           =   6135
      End
      Begin VB.Label remove_tenant_fr_label 
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
         Index           =   4
         Left            =   8400
         TabIndex        =   92
         Top             =   7080
         Width           =   1440
      End
      Begin VB.Label remove_tenant_fr_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   8400
         TabIndex        =   91
         Top             =   6240
         Width           =   525
      End
      Begin VB.Label remove_tenant_fr_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Remove Tenant"
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
         Left            =   10440
         TabIndex        =   86
         Top             =   3960
         Width           =   2250
      End
      Begin VB.Label remove_tenant_fr_label 
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
         TabIndex        =   85
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label remove_tenant_fr_goback 
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
         Left            =   11040
         TabIndex        =   84
         Top             =   8880
         Width           =   885
      End
      Begin VB.Label remove_tenant_fr_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Room"
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
         Left            =   8400
         TabIndex        =   83
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Image Image24 
         Height          =   1050
         Left            =   5520
         Picture         =   "Tenant Form.frx":218BA
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Image Image16 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   8160
         Picture         =   "Tenant Form.frx":2B5C4
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   6615
      End
      Begin VB.Image Image5 
         Height          =   615
         Left            =   8400
         Picture         =   "Tenant Form.frx":2BACE
         Stretch         =   -1  'True
         Top             =   8760
         Width           =   6135
      End
      Begin VB.Image Image4 
         Height          =   7215
         Left            =   6120
         Picture         =   "Tenant Form.frx":2BFD9
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   10695
      End
   End
   Begin VB.Frame tenant_frame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Tenant Detail Frame"
      Height          =   12375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   23055
      Begin VB.Frame Tenant_filter_frame 
         BorderStyle     =   0  'None
         Caption         =   "Tenant Filter Frame"
         Height          =   1335
         Left            =   15360
         TabIndex        =   95
         Top             =   10320
         Width           =   6855
         Begin VB.ComboBox tenant_filter_option_room 
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
            Left            =   2640
            TabIndex        =   102
            Text            =   "Room No."
            Top             =   600
            Width           =   1455
         End
         Begin VB.CommandButton tenant_filter_command 
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
            Height          =   495
            Left            =   4440
            TabIndex        =   96
            Top             =   720
            Width           =   2295
         End
         Begin VB.CheckBox tenant_filter_option 
            Caption         =   "Room number"
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
            Index           =   1
            Left            =   2400
            TabIndex        =   98
            Top             =   120
            Width           =   1815
         End
         Begin VB.CheckBox tenant_filter_option 
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
            TabIndex        =   97
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton tenant_filter_option_status 
            Caption         =   "Not living"
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
            Left            =   480
            TabIndex        =   100
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton tenant_filter_option_status 
            Caption         =   "Still living"
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
            Left            =   480
            TabIndex        =   99
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label tenant_filter_status 
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
            Left            =   4440
            TabIndex        =   101
            Top             =   240
            Width           =   2280
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H80000000&
            Height          =   1335
            Left            =   4320
            Top             =   0
            Width           =   2535
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H80000000&
            Height          =   1335
            Left            =   2160
            Top             =   0
            Width           =   2175
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000000&
            Height          =   1335
            Left            =   0
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   18480
         TabIndex        =   94
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   1440
         TabIndex        =   12
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   6615
         Left            =   2640
         TabIndex        =   10
         Top             =   3480
         Width           =   15
         Begin VB.Frame Frame4 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   6735
            Left            =   120
            TabIndex        =   11
            Top             =   2520
            Width           =   135
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   6960
         TabIndex        =   9
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   11280
         TabIndex        =   8
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   13080
         TabIndex        =   7
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   15840
         TabIndex        =   6
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   720
         TabIndex        =   5
         Top             =   3480
         Width           =   21495
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   720
         TabIndex        =   4
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   720
         TabIndex        =   3
         Top             =   10080
         Width           =   21495
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   720
         TabIndex        =   2
         Top             =   3960
         Width           =   21495
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   22200
         TabIndex        =   1
         Top             =   3480
         Width           =   15
      End
      Begin VB.TextBox tenant_fr_TextBox 
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
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "Tenant Form.frx":32725
         Top             =   4080
         Width           =   21375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Rent End Date"
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
         Left            =   18600
         TabIndex        =   93
         Top             =   3600
         Width           =   1560
      End
      Begin VB.Label tenant_fr_label_10 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Rent Start Date"
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
         Left            =   15960
         TabIndex        =   23
         Top             =   3600
         Width           =   1800
      End
      Begin VB.Label tenant_fr_label_7 
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
         Left            =   11400
         TabIndex        =   22
         Top             =   3600
         Width           =   1320
      End
      Begin VB.Label tenant_fr_label 
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
         TabIndex        =   21
         Top             =   2520
         Width           =   2355
      End
      Begin VB.Label tenant_fr_label_3 
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
         Left            =   840
         TabIndex        =   20
         Top             =   3600
         Width           =   480
      End
      Begin VB.Label tenant_fr_label 
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
         TabIndex        =   19
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label room_frame_main_label 
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
         TabIndex        =   18
         Top             =   11160
         Width           =   1035
      End
      Begin VB.Label tenant_fr_label_8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Citizenship No."
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
         Left            =   13200
         TabIndex        =   17
         Top             =   3600
         Width           =   1800
      End
      Begin VB.Label tenant_fr_label_6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Address(District/Municipality/Ward)"
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
         Left            =   7080
         TabIndex        =   16
         Top             =   3600
         Width           =   4200
      End
      Begin VB.Label tenant_fr_label_4 
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
         Left            =   1560
         TabIndex        =   15
         Top             =   3600
         Width           =   960
      End
      Begin VB.Label tenant_fr_label_5 
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
         Left            =   2760
         TabIndex        =   14
         Top             =   3600
         Width           =   1320
      End
      Begin VB.Image Image25 
         Height          =   1050
         Left            =   5520
         Picture         =   "Tenant Form.frx":327D1
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Image Image17 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   720
         Picture         =   "Tenant Form.frx":3C4DB
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Image Image26 
         Height          =   495
         Left            =   720
         Picture         =   "Tenant Form.frx":3C9E5
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   21495
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   720
         Picture         =   "Tenant Form.frx":3CEF0
         Stretch         =   -1  'True
         Top             =   11040
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Tenant_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim found As Boolean
Dim room_occupied As Boolean
Dim payment_status  As Boolean

Dim serial As Integer
Dim req_room_num As Integer
Dim counter_tenant_1 As Integer
Dim counter_tenant_2 As Integer
Dim counter_tenant_3 As Integer
Dim available_room_count As Integer

Dim temp_string As String
Dim temp_integer As Integer

Dim room_1 As room_class
Dim room_2 As room_class
Dim room_3 As room_class

Dim display_tenant As tenant_class_display

Dim new_tenant As tenant_class
Dim temp_tenant As tenant_class
Dim existing_tenant As tenant_class

Private Sub add_tenant_fr_Combo_1_GotFocus(Index As Integer)
    If Index = 1 Then
        If add_tenant_fr_Combo_1(1).Text = "Municipality" Then
            add_tenant_fr_Combo_1(1).Text = ""
        End If
    End If
End Sub

Private Sub add_tenant_fr_Combo_1_LostFocus(Index As Integer)
    If Index = 1 Then
        If add_tenant_fr_Combo_1(1).Text = "" Then
            add_tenant_fr_Combo_1(1).Text = "Municipality"
        End If
    End If
End Sub

Private Sub add_tenant_fr_combo_2_GotFocus()
    If add_tenant_fr_combo_2.Text = "Room number" Then add_tenant_fr_combo_2.Text = ""
End Sub

Private Sub add_tenant_fr_combo_2_LostFocus()
    If add_tenant_fr_combo_2.Text = "" Then add_tenant_fr_combo_2.Text = "Room number"
End Sub

Private Sub add_tenant_fr_combo_3_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 0 Then
            If add_tenant_fr_combo_3(0).Text <> "Year" Then add_tenant_fr_combo_3(1).SetFocus
        ElseIf Index = 1 Then
            If add_tenant_fr_combo_3(1).Text <> "Month" Then add_tenant_fr_add_command_Click
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub edit_room_frame_go_back_Click()
    'reset textbox values
    edit_tenant_fr_textbox_1(0).Text = "First Name"
    edit_tenant_fr_textbox_1(1).Text = "Middle Name"
    edit_tenant_fr_textbox_1(2).Text = "Last Name"
    edit_tenant_fr_textbox_1(3).Text = "Enter contact number"
    edit_tenant_fr_textbox_1(4).Text = "Enter citizenship number"
    
    edit_tenant_fr_combo_2(0).Text = "District"
    edit_tenant_fr_combo_2(1).Text = "Municipality"
    edit_tenant_fr_combo_2(2).Text = "Ward Number"
    
    If edit_tenant_fr_2.Visible = True Then
        edit_tenant_fr_1.Visible = True
        edit_tenant_fr_2.Visible = False
        
        'remove recenlty added room number for editing
        edit_tenant_fr_combo_1.Clear
        count_2 = room_detail_count_function
        Open "RoomDetail.txt" For Random As #1 Len = 119
            For count_1 = 1 To count_2 Step 1
                Get #1, count_1, room_1
                If room_1.room_occupied = False Then edit_tenant_fr_combo_1.AddItem room_1.room_number
            Next count_1
        Close #1
        
        edit_room_frame_save.Visible = False
        edit_tenant_fr_image_1.Visible = False
    Else
        Main_Form.WindowState = Tenant_Form.WindowState
        Main_Form.Visible = True
        Tenant_Form.Visible = False
    End If
End Sub

Private Sub edit_room_frame_save_Click()
    'check for empty room number in combo box
    If edit_tenant_fr_combo_1.Text = "" Or edit_tenant_fr_combo_1.Text = "Room Number" Then
        MsgBox_Response = MsgBox("     Please select the room number first.", vbInformation + vbOKOnly, "Rental Record")
        Exit Sub
    End If
    
    'check for default values
    If edit_tenant_fr_textbox_1(0).Text = "First Name" Or edit_tenant_fr_textbox_1(2).Text = "Last Name" Or edit_tenant_fr_textbox_1(2).Text = "Enter contact number" Or edit_tenant_fr_textbox_1(4).Text = "Enter citizenship number" Then
        If edit_tenant_fr_textbox_1(0).Text = "First Name" Then
            MsgBox_Response = MsgBox("     Please enter valid first name.", vbInformation + vbOKOnly, "Rental Record")
            edit_tenant_fr_textbox_1(0).SetFocus
        ElseIf edit_tenant_fr_textbox_1(2).Text = "Last Name" Then
            MsgBox_Response = MsgBox("     Please enter valid last name.", vbInformation + vbOKOnly, "Rental Record")
            edit_tenant_fr_textbox_1(2).SetFocus
        ElseIf edit_tenant_fr_textbox_1(3).Text = "Enter contact number" Then
            MsgBox_Response = MsgBox("     Please enter valid valid contact number.", vbInformation + vbOKOnly, "Rental Record")
            edit_tenant_fr_textbox_1(3).SetFocus
        ElseIf edit_tenant_fr_textbox_1(4).Text = "Enter citizenship number" Then
            MsgBox_Response = MsgBox("     Please enter valid citizenship number.", vbInformation + vbOKOnly, "Rental Record")
            edit_tenant_fr_textbox_1(4).SetFocus
        End If
        Exit Sub
    End If
    
    'check for default combo box values
    If edit_tenant_fr_combo_2(0).Text = "District" Or edit_tenant_fr_combo_2(1).Text = "Municipality" Or edit_tenant_fr_combo_2(2).Text = "Ward Number" Then
        If edit_tenant_fr_combo_2(0).Text = "District" Then
            MsgBox_Response = MsgBox("     Please select the valid district.", vbInformation + vbOKOnly, "Rental Record")
            edit_tenant_fr_combo_2(0).SetFocus
        ElseIf edit_tenant_fr_combo_2(1).Text = "" Then
            MsgBox_Response = MsgBox("     Please select the valid municipality.", vbInformation + vbOKOnly, "Rental Record")
            edit_tenant_fr_combo_2(1).SetFocus
        ElseIf edit_tenant_fr_combo_2(2).Text = "" Then
            MsgBox_Response = MsgBox("     Please select the valid ward number.", vbInformation + vbOKOnly, "Rental Record")
            edit_tenant_fr_combo_2(2).SetFocus
        End If
        Exit Sub
    End If
    
    'check for municipality validity
    found = False
    For i = 0 To edit_tenant_fr_combo_2(1).ListCount Step 1
        If edit_tenant_fr_combo_2(1).Text = edit_tenant_fr_combo_2(1).List(i) Then
            found = True
        End If
    Next i
    
    If found = False Then
        MsgBox_Response = MsgBox("          Please select the valid municipality.", vbInformation + vbOKOnly, "Rental Record")
        edit_tenant_fr_combo_2(1).SetFocus
        Exit Sub
    End If
    
    'copy new values
    temp_tenant.room_num = Val(edit_tenant_fr_combo_1.Text)
    temp_tenant.first_name = edit_tenant_fr_textbox_1(0).Text
    
    If edit_tenant_fr_textbox_1(1).Text = "Middle Name" Then
        temp_tenant.middle_name = "Null"
    Else
        temp_tenant.middle_name = edit_tenant_fr_textbox_1(1).Text
    End If
    
    temp_tenant.last_name = edit_tenant_fr_textbox_1(2).Text
    temp_tenant.address_district = edit_tenant_fr_combo_2(0).Text
    temp_tenant.address_municipality = edit_tenant_fr_combo_2(1).Text
    temp_tenant.address_ward = edit_tenant_fr_combo_2(2).Text
    temp_tenant.contact_num = edit_tenant_fr_textbox_1(3).Text
    temp_tenant.citizenship = edit_tenant_fr_textbox_1(4).Text
    
    room_1.room_number = Val(edit_tenant_fr_label(9).Caption)
    
    'for no room change
    If room_1.room_number = Val(edit_tenant_fr_combo_1.Text) Then
        'replace tenant detail in a tenant detail file
        counter_tenant_2 = tenant_detail_count_function
        
        Open "TenantDetail.txt" For Random As #1 Len = 103
            For counter_tenant_1 = 1 To counter_tenant_2 Step 1
                Get #1, counter_tenant_1, existing_tenant
                If existing_tenant.room_num = room_1.room_number Then
                    temp_tenant.rent_year(0) = existing_tenant.rent_year(0)
                    temp_tenant.rent_year(1) = existing_tenant.rent_year(1)
                    temp_tenant.rent_month(0) = existing_tenant.rent_month(0)
                    temp_tenant.rent_month(1) = existing_tenant.rent_month(1)
                    Put #1, counter_tenant_1, temp_tenant
                End If
            Next counter_tenant_1
        Close #1
        
        'change tenant detail from rent detail file
        counter_tenant_2 = room_detail_count_function
        Open "RoomDetail.txt" For Random As #2 Len = 112
            For counter_tenant_1 = 1 To counter_tenant_2 Step 1
                Get #2, counter_tenant_1, room_2
                
                If room_2.room_number = Val(edit_tenant_fr_label(9).Caption) Then
                    room_2.tenant_fname = edit_tenant_fr_textbox_1(0).Text
                    
                    If edit_tenant_fr_textbox_1(1).Text = "Middle Name" Then
                        room_2.tenant_mname = "Null"
                    Else
                        room_2.tenant_mname = edit_tenant_fr_textbox_1(1).Text
                    End If
                    
                    room_2.tenant_lname = edit_tenant_fr_textbox_1(2).Text
                    room_2.tenant_contact = edit_tenant_fr_textbox_1(3).Text
                    
                    Put #2, counter_tenant_1, room_2
                End If
            Next counter_tenant_1
        Close #2
        
        MsgBox_Response = MsgBox("     Tenant details have been modified successfully. Do you want to edit another detail?", vbInformation + vbYesNo, "Details modified successfully")
        edit_room_frame_go_back_Click
    End If
End Sub

'proceed command button press
Private Sub edit_tenant_fr_1_Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'proceed button press
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub edit_tenant_fr_1_command_1_Click()
    req_room_num = Val(edit_tenant_fr_1_Combo1.Text)
    If edit_tenant_fr_1_Combo1.Text = "" Or edit_tenant_fr_1_Combo1.Text = "Room number" Then
        MsgBox_Response = MsgBox("          Please select the room number first.", vbInformation + vbOKOnly, "Rental Record")
    Else
        edit_tenant_fr_combo_1.AddItem edit_tenant_fr_1_Combo1.Text
        
        counter_tenant_2 = tenant_detail_count_function
        
        Open "TenantDetail.txt" For Random As #1 Len = 103
            'load tenant details in a tenant file
            For counter_tenant_1 = 1 To counter_tenant_2 Step 1
                Get #1, counter_tenant_1, existing_tenant
                If existing_tenant.room_num = req_room_num Then
                    temp_string = "Null      "
                    
                    If existing_tenant.middle_name = temp_string Then 'has no middle name
                        edit_tenant_fr_label(10).Caption = existing_tenant.first_name + existing_tenant.last_name
                    Else
                        edit_tenant_fr_label(10).Caption = existing_tenant.first_name + existing_tenant.middle_name + existing_tenant.last_name
                    End If
                    
                    edit_tenant_fr_label(9).Caption = existing_tenant.room_num
                    edit_tenant_fr_label(12).Caption = existing_tenant.address_district
                    edit_tenant_fr_label(14).Caption = existing_tenant.address_municipality
                    edit_tenant_fr_label(16).Caption = existing_tenant.address_ward
                    edit_tenant_fr_label(17).Caption = existing_tenant.contact_num
                    edit_tenant_fr_label(18).Caption = existing_tenant.citizenship
                    
                    edit_tenant_fr_combo_2(0) = edit_tenant_fr_label(12).Caption
                    edit_tenant_fr_combo_2(1) = edit_tenant_fr_label(14).Caption
                    edit_tenant_fr_combo_2(2) = edit_tenant_fr_label(16).Caption
                    
                    
                    edit_tenant_fr_2.Visible = True
                    edit_tenant_fr_1.Visible = False

                    edit_room_frame_save.Visible = True
                    edit_tenant_fr_image_1.Visible = True
                    
                    edit_tenant_fr_combo_1.Clear
                    edit_tenant_fr_combo_1.AddItem req_room_num
                    edit_tenant_fr_combo_1.Text = "Room number"
                End If
            Next counter_tenant_1
        Close #1
    End If
End Sub

Private Sub edit_tenant_fr_combo_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If edit_tenant_fr_combo_1.Text <> "" And edit_tenant_fr_combo_1.Text <> "Room Number" Then
            KeyAscii = 0
            edit_tenant_fr_textbox_1(0).SetFocus
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub edit_tenant_fr_combo_2_GotFocus(Index As Integer)
    If Index = 0 Then
        If edit_tenant_fr_combo_2(0).Text = "District" Then edit_tenant_fr_combo_2(0).Text = ""
    ElseIf Index = 1 Then
        If edit_tenant_fr_combo_2(1).Text = "Municipality" Then edit_tenant_fr_combo_2(1).Text = ""
    ElseIf Index = 2 Then
        If edit_tenant_fr_combo_2(2).Text = "Ward Number" Then edit_tenant_fr_combo_2(2).Text = ""
    End If
End Sub

Private Sub edit_tenant_fr_combo_2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 0 Then
            If edit_tenant_fr_combo_2(0).Text <> "" And edit_tenant_fr_combo_2(0).Text <> "District" Then edit_tenant_fr_combo_2(1).SetFocus
        ElseIf Index = 1 Then
            If edit_tenant_fr_combo_2(1).Text <> "" And edit_tenant_fr_combo_2(1).Text <> "Municipality" Then edit_tenant_fr_combo_2(2).SetFocus
        ElseIf Index = 2 Then
            If edit_tenant_fr_combo_2(2).Text <> "" And edit_tenant_fr_combo_2(2).Text <> "Ward Number" Then edit_tenant_fr_textbox_1(3).SetFocus
        End If
    Else
        If KeyAscii = 8 Or KeyAscii = 32 Then
            KeyAscii = 0
        ElseIf KeyAscii >= 48 And keyasci <= 57 Then
            KeyAscii = 0
        Else
            kwyascii = KeyAscii
        End If
    End If
End Sub

Private Sub edit_tenant_fr_combo_2_LostFocus(Index As Integer)
    If Index = 0 Then
        If edit_tenant_fr_combo_2(0).Text = "" Then edit_tenant_fr_combo_2(0).Text = "District"
    ElseIf Index = 1 Then
        If edit_tenant_fr_combo_2(1).Text = "" Then edit_tenant_fr_combo_2(1).Text = "Municipality"
    ElseIf Index = 2 Then
        If edit_tenant_fr_combo_2(2).Text = "" Then edit_tenant_fr_combo_2(2).Text = "Ward Number"
    End If
End Sub


Private Sub edit_tenant_fr_textbox_1_GotFocus(Index As Integer)
    If Index = 0 Then
        If edit_tenant_fr_textbox_1(0).Text = "First Name" Then edit_tenant_fr_textbox_1(0).Text = ""
    ElseIf Index = 1 Then
        If edit_tenant_fr_textbox_1(1).Text = "Middle Name" Then edit_tenant_fr_textbox_1(1).Text = ""
    ElseIf Index = 2 Then
        If edit_tenant_fr_textbox_1(2).Text = "Last Name" Then edit_tenant_fr_textbox_1(2).Text = ""
    ElseIf Index = 3 Then
        If edit_tenant_fr_textbox_1(3).Text = "Enter contact number" Then edit_tenant_fr_textbox_1(3).Text = ""
    ElseIf Index = 4 Then
        If edit_tenant_fr_textbox_1(4).Text = "Enter citizenship number" Then edit_tenant_fr_textbox_1(4).Text = ""
    End If
End Sub

Private Sub edit_tenant_fr_textbox_1_LostFocus(Index As Integer)
    If Index = 0 Then
        If edit_tenant_fr_textbox_1(0).Text = "" Then edit_tenant_fr_textbox_1(0).Text = "First Name"
    ElseIf Index = 1 Then
        If edit_tenant_fr_textbox_1(1).Text = "" Then edit_tenant_fr_textbox_1(1).Text = "Middle Name"
    ElseIf Index = 2 Then
        If edit_tenant_fr_textbox_1(2).Text = "" Then edit_tenant_fr_textbox_1(2).Text = "Last Name"
    ElseIf Index = 3 Then
        If edit_tenant_fr_textbox_1(3).Text = "" Then edit_tenant_fr_textbox_1(3).Text = "Enter contact number"
    ElseIf Index = 4 Then
        If edit_tenant_fr_textbox_1(4).Text = "" Then edit_tenant_fr_textbox_1(4).Text = "Enter citizenship number"
    End If
End Sub

Private Sub edit_tenant_fr_textbox_1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 0 Then
            If edit_tenant_fr_textbox_1(0).Text <> "" Then edit_tenant_fr_textbox_1(1).SetFocus
        ElseIf Index = 1 Then
            If edit_tenant_fr_textbox_1(1).Text = "" Then
                edit_tenant_fr_textbox_1(1).Text = "Middle Name"
                edit_tenant_fr_textbox_1(2).SetFocus
            ElseIf edit_tenant_fr_textbox_1(1).Text <> "" Then
                edit_tenant_fr_textbox_1(2).SetFocus
            End If
            'If edit_tenant_fr_textbox_1(1).Text <> "" Then edit_tenant_fr_textbox_1(2).SetFocus
        ElseIf Index = 2 Then
            If edit_tenant_fr_textbox_1(2).Text <> "" Then edit_tenant_fr_combo_2(0).SetFocus
        ElseIf Index = 3 Then
            If edit_tenant_fr_textbox_1(3).Text <> "" Then edit_tenant_fr_textbox_1(4).SetFocus
        ElseIf Index = 4 Then
            If edit_tenant_fr_textbox_1(4).Text <> "" Then
                'save button press
            End If
        End If
        Exit Sub
    ElseIf KeyAscii = 32 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If Index = 3 Or Index = 4 Then
        If KeyAscii = 8 Then
            KeyAscii = KeyAscii
        ElseIf KeyAscii = 45 And Index = 4 Then
            KeyAscii = KeyAscii
        ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
            KeyAscii = 0
        End If
    End If
End Sub


Private Sub remove_tenant_fr_combo_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And remove_tenant_fr_combo_1.Text <> "" Then
        remove_tenant_fr_TextBox(0).SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub remove_tenant_fr_remove_command_Click()
    req_room_num = Val(remove_tenant_fr_combo_1.Text)
    
    If remove_tenant_fr_combo_1.Text = "" Then
        MsgBox_Response = MsgBox("     Please select the room number first.", vbInformation + vbOKOnly, "Rental Record")
        
    ElseIf remove_tenant_fr_combo_1.Text <> "" Then
        If remove_tenant_fr_remove_command.Caption = "Remove" Then
            counter_tenant_2 = room_detail_count_function
            Open "RoomDetail.txt" For Random As #1 Len = 112
                For counter_tenant_1 = 1 To counter_tenant_2 Step 1
                    Get #1, counter_tenant_1, room_1
                    If room_1.room_number = req_room_num And room_1.room_occupied = True Then
                        remove_tenant_fr_TextBox(0).Text = room_1.tenant_fname
                        
                        If room_1.tenant_mname = "Null      " Then
                            remove_tenant_fr_TextBox(1).Text = "Middle Name"
                        Else
                            remove_tenant_fr_TextBox(1).Text = room_1.tenant_mname
                        End If
                        
                        remove_tenant_fr_TextBox(2).Text = room_1.tenant_lname
                    End If
                Next counter_tenant_1
            Close #1
            remove_tenant_fr_TextBox(3).SetFocus
            remove_tenant_fr_remove_command.Caption = "Confirm"
            
        ElseIf remove_tenant_fr_remove_command.Caption = "Confirm" Then
            If remove_tenant_fr_TextBox(3).Text = "Enter contact number" Then 'has not entered contact number
                MsgBox_Response = MsgBox("     Please enter the contact number.", vbInformation + vbOKOnly, "Rental Record")
            ElseIf remove_tenant_fr_TextBox(3).Text <> "" Then 'has entered contact number
                'check if the contact number is valid
                found = False
                Open "RoomDetail.txt" For Random As #2 Len = 112
                    For counter_tenant_1 = 1 To counter_tenant_2 Step 1
                        Get #2, counter_tenant_1, room_1
                        If room_1.room_number = req_room_num And room_1.room_occupied = True Then
                            room_2.tenant_contact = remove_tenant_fr_TextBox(3)
                            If room_1.tenant_contact = room_2.tenant_contact Then
                                found = True
                                counter_tenant_1 = counter_tenant_2 + 1
                            End If
                        End If
                    Next counter_tenant_1
                Close #2
                
                If found = False Then
                    MsgBox_Response = MsgBox("     Make sure you entered the contact number correctly.", vbInformation + vbOKOnly, "Rental Record")
                Else
                    'check if the tenant has rent left to be paid
                    room_1.room_number = Val(remove_tenant_fr_combo_1.Text)
                    room_1.tenant_fname = remove_tenant_fr_TextBox(0)
                    
                    If remove_tenant_fr_TextBox(1).Text = "Middle Name" Then
                        room_1.tenant_mname = "Null"
                    Else
                        room_1.tenant_mname = remove_tenant_fr_TextBox(1)
                    End If
                    
                    room_1.tenant_lname = remove_tenant_fr_TextBox(2)
                                            
                    payment_status = False
                    
                    payment_status = pending_payment_check(room_1.room_number, room_1.tenant_fname, room_1.tenant_mname, room_1.tenant_lname)
                
                    If payment_status = False Then 'has no rent left to be paid
                        'remove tenant detail from room detail file
                        counter_tenant_2 = room_detail_count_function
                        Open "RoomDetail.txt" For Random As #3 Len = 112
                            For counter_tenant_1 = 1 To counter_tenant_2 Step 1
                                Get #3, counter_tenant_1, room_1
                                If req_room_num = room_1.room_number And room_1.room_occupied = True Then
                                    room_1.room_occupied = False
                                    room_1.tenant_fname = "Null"
                                    room_1.tenant_mname = "Null"
                                    room_1.tenant_lname = "Null"
                                    room_1.tenant_contact = "Null"
                                    room_1.service_provided = "Null"
                                    Put #3, counter_tenant_1, room_1
                                End If
                            Next counter_tenant_1
                        Close #3
                        
                        'place current date as rent end date in tenant detail file
                        counter_tenant_2 = tenant_detail_count_function
                        Open "TenantDetail.txt" For Random As #4 Len = 103
                            For counter_tenant_1 = 1 To counter_tenant_2 Step 1
                                Get #4, counter_tenant_1, existing_tenant
                                If existing_tenant.room_num = req_room_num Then
                                    temp_tenant.contact_num = remove_tenant_fr_TextBox(3).Text
                                    If existing_tenant.contact_num = temp_tenant.contact_num Then
                                        existing_tenant.rent_year(1) = return_year
                                        existing_tenant.rent_month(1) = return_month
                                        Put #4, counter_tenant_1, existing_tenant
                                    End If
                                End If
                            Next counter_tenant_1
                        Close #4

                        MsgBox_Response = MsgBox("          Tenant removed successfully.", vbInformation + vbOKOnly, "Rental Record")
                        remove_tenant_fr_goback_Click
                    Else 'rent left to be paid
                        MsgBox_Response = MsgBox("   Sorry! this tenant has rent left to be paid.", vbInformation + vbOKOnly, "Rental Record")
                        Call show_pending_payment(room_1.room_number, room_1.tenant_fname, room_1.tenant_mname, room_1.tenant_lname)
                    End If
                End If
            End If
        End If
    End If
End Sub

'go back label
Private Sub room_frame_main_label_Click()
    tenant_filter_option(0).Value = Unchecked
    tenant_filter_option(1).Value = Unchecked
    tenant_filter_option_room.Text = ""
    tenant_filter_status.Caption = "Filter status : Off"
    
    Main_Form.WindowState = Tenant_Form.WindowState
    Main_Form.Visible = True
    Tenant_Form.Visible = False
    tenant_frame.Visible = False
End Sub

'add command button
Private Sub add_tenant_fr_add_command_Click()
    'check for empty room number
    If add_tenant_fr_combo_2.Text = "" Then
        MsgBox_Response = MsgBox("          Please select the room number first.", vbInformation + vbOKOnly, "Rental Record")
        add_tenant_fr_combo_2.SetFocus
        Exit Sub
    Else
        'check for default combo box values
        If add_tenant_fr_combo_2.Text = "Room number" Or add_tenant_fr_Combo_1(0).Text = "District" Or add_tenant_fr_Combo_1(1).Text = "Municipality" Or add_tenant_fr_Combo_1(2).Text = "Ward" Then
            If add_tenant_fr_combo_2.Text = "Room number" Then
                MsgBox_Response = MsgBox("          Please choose valid room number.", vbInformation + vbOKOnly, "Rental Record")
                add_tenant_fr_combo_2.SetFocus
            ElseIf add_tenant_fr_Combo_1(0).Text = "District" Then
                MsgBox_Response = MsgBox("          Please choose valid district.", vbInformation + vbOKOnly, "Rental Record")
                add_tenant_fr_Combo_1(0).SetFocus
            ElseIf add_tenant_fr_Combo_1(1).Text = "Municipality" Then
                MsgBox_Response = MsgBox("          Please choose valid municipality.", vbInformation + vbOKOnly, "Rental Record")
                add_tenant_fr_Combo_1(1).SetFocus
            ElseIf add_tenant_fr_Combo_1(2).Text = "Ward" Then
                MsgBox_Response = MsgBox("          Please enter valid ward number.", vbInformation + vbOKOnly, "Rental Record")
                add_tenant_fr_Combo_1(2).SetFocus
            End If
            Exit Sub
        End If
        
        'check for municipality validity
        found = False
        For i = 0 To add_tenant_fr_Combo_1(1).ListCount Step 1
            Debug.Print add_tenant_fr_Combo_1(1).List(i)
            If add_tenant_fr_Combo_1(1).Text = add_tenant_fr_Combo_1(1).List(i) Then
                found = True
            End If
        Next i
        
        If found = False Then
            MsgBox_Response = MsgBox("          Please select the valid municipality.", vbInformation + vbOKOnly, "Rental Record")
            add_tenant_fr_Combo_1(1).SetFocus
            Exit Sub
        End If
        
        'default textbox values test
        If add_tenant_fr_textbox(0).Text = "First Name" Or add_tenant_fr_textbox(2).Text = "Last Name" Or add_tenant_fr_textbox(3).Text = "Enter contact number" Or add_tenant_fr_textbox(4).Text = "Enter citizenship number" Then
            If add_tenant_fr_textbox(0).Text = "First Name" Then
                MsgBox_Response = MsgBox("          Please enter valid first name.", vbInformation + vbOKOnly, "Rental Record")
                add_tenant_fr_textbox(0).SetFocus
            ElseIf add_tenant_fr_textbox(2).Text = "Last Name" Then
                MsgBox_Response = MsgBox("          Please enter valid last name.", vbInformation + vbOKOnly, "Rental Record")
                add_tenant_fr_textbox(2).SetFocus
            ElseIf add_tenant_fr_textbox(3).Text = "Enter contact number" Then
                MsgBox_Response = MsgBox("          Please enter valid contact number.", vbInformation + vbOKOnly, "Rental Record")
                add_tenant_fr_textbox(3).SetFocus
            ElseIf add_tenant_fr_textbox(4).Text = "Enter citizenship number" Then
                MsgBox_Response = MsgBox("          Please enter valid citizenship number.", vbInformation + vbOKOnly, "Rental Record")
                add_tenant_fr_textbox(4).SetFocus
            End If
            Exit Sub
        End If
        
        'default combo values >> rent date
        If add_tenant_fr_combo_3(0).Text = "Year" Or add_tenant_fr_combo_3(1).Text = "Month" Then
            If add_tenant_fr_combo_3(0).Text = "Year" Then
                MsgBox_Response = MsgBox("          Please enter valid rent date - year.", vbInformation + vbOKOnly, "Rental Record")
                add_tenant_fr_combo_3(0).SetFocus
            Else
                MsgBox_Response = MsgBox("          Please enter valid rent date - month.", vbInformation + vbOKOnly, "Rental Record")
                add_tenant_fr_combo_3(1).SetFocus
            End If
            Exit Sub
        End If
        
        'assign values
        new_tenant.first_name = add_tenant_fr_textbox(0).Text
            
        If add_tenant_fr_textbox(1).Text = "Middle Name" Then
            new_tenant.middle_name = "Null"
        Else
            new_tenant.middle_name = add_tenant_fr_textbox(1).Text
        End If
            
        new_tenant.last_name = add_tenant_fr_textbox(2).Text
        new_tenant.address_district = add_tenant_fr_Combo_1(0).Text
        new_tenant.address_municipality = add_tenant_fr_Combo_1(1).Text
        new_tenant.address_ward = add_tenant_fr_Combo_1(2).Text
        new_tenant.citizenship = add_tenant_fr_textbox(4).Text
        new_tenant.contact_num = add_tenant_fr_textbox(3).Text
        new_tenant.room_num = Val(add_tenant_fr_combo_2.Text)
        
        new_tenant.rent_year(0) = Val(add_tenant_fr_combo_3(0).Text)
        new_tenant.rent_year(1) = 0
        new_tenant.rent_month(1) = 0
        
        new_tenant.rent_month(0) = set_month_in_integer(add_tenant_fr_combo_3(1).Text)

        'check if the provided date is in the past
        counter_tenant_2 = return_year
        counter_tenant_3 = return_month
        
        If new_tenant.rent_year(0) = counter_tenant_2 Then
            If new_tenant.rent_month(0) < counter_tenant_3 Then
                MsgBox_Response = MsgBox("     You cannot add tenant for past date.", vbInformation + vbOKOnly, "Rental Record")
                Exit Sub
            End If
        End If
        
            
        'write tenant detail in a file
        counter_tenant_1 = tenant_detail_count_function
        
        Open "TenantDetail.txt" For Random As #2 Len = 103
            Put #2, counter_tenant_1 + 1, new_tenant
        Close #2
        
        counter_tenant_2 = room_detail_count_function
        
        req_room_num = Val(add_tenant_fr_combo_2.Text)
        
        Open "RoomDetail.txt" For Random As #3 Len = 112
            For counter_tenant_1 = 1 To counter_tenant_2 Step 1
                Get #3, counter_tenant_1, room_1
                If room_1.room_number = req_room_num Then
                    room_1.room_occupied = True
                    room_1.tenant_fname = add_tenant_fr_textbox(0).Text
                    
                    If add_tenant_fr_textbox(1).Text = "Middle Name" Then
                        room_1.tenant_mname = "Null"
                    Else
                        room_1.tenant_mname = add_tenant_fr_textbox(1).Text
                    End If
                    
                    room_1.tenant_lname = add_tenant_fr_textbox(2).Text
                    room_1.tenant_contact = add_tenant_fr_textbox(3).Text
                    
                    If add_tenant_fr_check(4).Value = Unchecked Then
                        room_1.service_provided = "Water, Waste Management, Electricity, Security"
                    Else
                        room_1.service_provided = "Water, Waste Management, Electricity, Security, Internet"
                    End If
                    
                    Put #3, counter_tenant_1, room_1
                End If
            Next counter_tenant_1
        Close #3
        
        MsgBox_Response = MsgBox("     Do you want to add another tenant detail?", vbInformation + vbYesNo, "Tenant detail added successfully!")
        
        'reset user provided values
        add_tenant_fr_combo_2.Text = "Room number"
        add_tenant_fr_Combo_1(0).Text = "District"
        add_tenant_fr_Combo_1(1).Text = "Municipality"
        add_tenant_fr_Combo_1(2).Text = "Ward"
        add_tenant_fr_textbox(0).Text = "First Name"
        add_tenant_fr_textbox(1).Text = "Middle Name"
        add_tenant_fr_textbox(2).Text = "Last Name"
        add_tenant_fr_textbox(3).Text = "Enter contact number"
        add_tenant_fr_textbox(4).Text = "Enter citizenship number"
        add_tenant_fr_combo_3(0).Text = "Year"
        add_tenant_fr_combo_3(1).Text = "Month"
        add_tenant_fr_check(4).Value = Unchecked
        
        If MsgBox_Response = 6 Then
            counter_tenant_2 = room_detail_count_function
            add_tenant_fr_combo_2.Clear
            Open "RoomDetail.txt" For Random As #1 Len = 112
                For counter_tenant_1 = 1 To counter_tenant_2 Step 1
                    Get #1, counter_tenant_1, room_1
                    If room_1.room_occupied = False Then
                        add_tenant_fr_combo_2.AddItem room_1.room_number
                    End If
                Next counter_tenant_1
            Close #1
            add_tenant_fr_combo_2.SetFocus
        Else
            add_tenant_fr_goback_Click
        End If
    End If
End Sub
'go back label
Private Sub add_tenant_fr_goback_Click()
    Main_Form.WindowState = Tenant_Form.WindowState
    Main_Form.Visible = True
    Tenant_Form.Visible = False
    Tenant_Form.add_tenant_frame.Visible = False
End Sub

Private Sub add_tenant_fr_combo_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If add_tenant_fr_combo_2.Text <> "" And add_tenant_fr_combo_2.Text <> "Room number" Then
            KeyAscii = 0
            add_tenant_fr_textbox(0).SetFocus
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub add_tenant_fr_textbox_GotFocus(Index As Integer)
    If Index = 0 Then
        If add_tenant_fr_textbox(0).Text = "First Name" Then add_tenant_fr_textbox(0).Text = ""
    ElseIf Index = 1 Then
        If add_tenant_fr_textbox(1).Text = "Middle Name" Then add_tenant_fr_textbox(1).Text = ""
    ElseIf Index = 2 Then
        If add_tenant_fr_textbox(2).Text = "Last Name" Then add_tenant_fr_textbox(2).Text = ""
    ElseIf Index = 3 Then
        If add_tenant_fr_textbox(3).Text = "Enter contact number" Then add_tenant_fr_textbox(3).Text = ""
    ElseIf Index = 4 Then
        If add_tenant_fr_textbox(4).Text = "Enter citizenship number" Then add_tenant_fr_textbox(4).Text = ""
    End If
End Sub

Private Sub add_tenant_fr_textbox_KeyPress(Index As Integer, KeyAscii As Integer)
    'enter key press
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 0 Then
            If add_tenant_fr_textbox(0).Text <> "" Then add_tenant_fr_textbox(1).SetFocus
        ElseIf Index = 1 Then
            add_tenant_fr_textbox(2).SetFocus
        ElseIf Index = 2 Then
            If add_tenant_fr_textbox(2).Text <> "" Then add_tenant_fr_Combo_1(0).SetFocus
        ElseIf Index = 3 Then
            If add_tenant_fr_textbox(3).Text <> "" Then add_tenant_fr_textbox(4).SetFocus
        ElseIf Index = 4 Then
            If add_tenant_fr_textbox(4).Text <> "" Then add_tenant_fr_combo_3(0).SetFocus
        End If
        Exit Sub
    End If
    
    If KeyAscii = 32 Or KeyAscii = 8 Then 'space key press
        If KeyAscii = 32 Then
            KeyAscii = 0
        ElseIf KeyAscii = 8 Then
            KeyAscii = KeyAscii
        End If
        Exit Sub
    End If
    
    If KeyAscii = 45 And Index = 4 Then
        KeyAscii = KeyAscii
        Exit Sub
    End If
    
    If KeyAscii <> 13 And KeyAscii <> 32 Then
        If Index = 3 Or Index = 4 Then
            If KeyAscii < 47 Or KeyAscii > 58 Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub add_tenant_fr_textbox_LostFocus(Index As Integer)
    If Index = 0 Then
        If add_tenant_fr_textbox(0).Text = "" Then add_tenant_fr_textbox(0).Text = "First Name"
    ElseIf Index = 1 Then
        If add_tenant_fr_textbox(1).Text = "" Then add_tenant_fr_textbox(1).Text = "Middle Name"
    ElseIf Index = 2 Then
        If add_tenant_fr_textbox(2).Text = "" Then add_tenant_fr_textbox(2).Text = "Last Name"
    ElseIf Index = 3 Then
        If add_tenant_fr_textbox(3).Text = "" Then add_tenant_fr_textbox(3).Text = "Enter contact number"
    ElseIf Index = 4 Then
        If add_tenant_fr_textbox(4).Text = "" Then add_tenant_fr_textbox(4).Text = "Enter citizenship number"
    End If
End Sub

Private Sub add_tenant_fr_Combo_1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 0 Then
            If add_tenant_fr_Combo_1(0).Text <> "" And add_tenant_fr_Combo_1(0).Text <> "District" Then add_tenant_fr_Combo_1(1).SetFocus
        ElseIf Index = 1 Then
            If add_tenant_fr_Combo_1(1).Text <> "" And add_tenant_fr_Combo_1(1).Text <> "Municipality" Then add_tenant_fr_Combo_1(2).SetFocus
        ElseIf Index = 2 Then
            If add_tenant_fr_Combo_1(2).Text <> "" And add_tenant_fr_Combo_1(2).Text <> "Ward" Then add_tenant_fr_textbox(3).SetFocus
        End If
    Else
        If KeyAscii = 8 Or KeyAscii = 32 Then
            KeyAscii = KeyAscii
        ElseIf KeyAscii >= 48 And KeyAscii <= 57 Then
            KeyAscii = 0
        Else
            KeyAscii = KeyAscii
        End If
        'KeyAscii = 0
        'check for validity of input
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Remove Tenant Frame
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub remove_tenant_fr_goback_Click()
    'reset user provided values
    remove_tenant_fr_combo_1.Clear
    remove_tenant_fr_TextBox(0).Text = "First Name"
    remove_tenant_fr_TextBox(1).Text = "Middle Name"
    remove_tenant_fr_TextBox(2).Text = "Last Name"
    remove_tenant_fr_TextBox(3).Text = "Enter contact number"
    remove_tenant_fr_remove_command.Caption = "Remove"
    
    'frame visibility
    Main_Form.WindowState = Tenant_Form.WindowState
    Main_Form.Visible = True
    Tenant_Form.Visible = False
    Tenant_Form.remove_tenant_frame.Visible = False
End Sub

Private Sub remove_tenant_fr_TextBox_GotFocus(Index As Integer)
    If Index = 0 Then
        If remove_tenant_fr_TextBox(0).Text = "First Name" Then remove_tenant_fr_TextBox(0).Text = ""
    ElseIf Index = 1 Then
        If remove_tenant_fr_TextBox(1).Text = "Middle Name" Then remove_tenant_fr_TextBox(1).Text = ""
    ElseIf Index = 2 Then
        If remove_tenant_fr_TextBox(2).Text = "Last Name" Then remove_tenant_fr_TextBox(2).Text = ""
    ElseIf Index = 3 Then
        If remove_tenant_fr_TextBox(3).Text = "Enter contact number" Then remove_tenant_fr_TextBox(3).Text = ""
    End If
End Sub

Private Sub remove_tenant_fr_TextBox_LostFocus(Index As Integer)
    If Index = 0 Then
        If remove_tenant_fr_TextBox(0).Text = "" Then remove_tenant_fr_TextBox(0).Text = "First Name"
    ElseIf Index = 1 Then
        If remove_tenant_fr_TextBox(1).Text = "" Then remove_tenant_fr_TextBox(1).Text = "Middle Name"
    ElseIf Index = 2 Then
        If remove_tenant_fr_TextBox(2).Text = "" Then remove_tenant_fr_TextBox(2).Text = "Last Name"
    ElseIf Index = 3 Then
        If remove_tenant_fr_TextBox(3).Text = "" Then remove_tenant_fr_TextBox(3).Text = "Enter contact number"
    End If
End Sub

Private Sub remove_tenant_fr_TextBox_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 0 Then
            If remove_tenant_fr_TextBox(0).Text <> "" Then remove_tenant_fr_TextBox(1).SetFocus
        ElseIf Index = 1 Then
            If remove_tenant_fr_TextBox(1).Text <> "" Then remove_tenant_fr_TextBox(2).SetFocus
        ElseIf Index = 2 Then
            If remove_tenant_fr_TextBox(2).Text <> "" Then remove_tenant_fr_TextBox(3).SetFocus
        ElseIf Index = 3 Then
            If remove_tenant_fr_TextBox(3).Text <> "" Then remove_tenant_fr_remove_command_Click
        End If
    End If
End Sub


'FILTER TENANT DETAILS
Private Sub tenant_filter_command_Click()
    If tenant_filter_option(0).Value = Unchecked And tenant_filter_option(1).Value = Unchecked Then 'filter : off
        tenant_filter_status.Caption = "Filter status : Off"

        tenant_fr_TextBox.Text = ""
        counter_tenant_2 = tenant_detail_count_function
        
        If counter_tenant_2 = 0 Then
            tenant_fr_TextBox.Text = "No detail has been added yet!"
        Else
            Open "TenantDetail.txt" For Random As #1 Len = 103
                For counter_tenant_1 = 1 To counter_tenant_2 Step 1
                    Get #1, counter_tenant_1, existing_tenant
                    
                    display_tenant.serial = CStr(counter_tenant_1)
                    display_tenant.room_num = CStr(existing_tenant.room_num)
                    display_tenant.first_name = existing_tenant.first_name
                    display_tenant.address_district = existing_tenant.address_district
                    display_tenant.address_municipality = existing_tenant.address_municipality
                    display_tenant.address_ward = existing_tenant.address_ward
                    display_tenant.citizenship = existing_tenant.citizenship
                    display_tenant.contact_num = existing_tenant.contact_num
                    
                    display_tenant.rent_year(0) = CStr(existing_tenant.rent_year(0))
                    display_tenant.rent_year(1) = CStr(existing_tenant.rent_year(1))
                    
                    Call return_month_in_string(display_tenant.rent_month(0), existing_tenant.rent_month(0))
                    Call return_month_in_string(display_tenant.rent_month(1), existing_tenant.rent_month(1))
                
                    If existing_tenant.middle_name = "Null      " Then ' has no middle name
                        display_tenant.middle_name = existing_tenant.last_name
                        display_tenant.last_name = ""
                    Else
                        display_tenant.middle_name = existing_tenant.middle_name
                        display_tenant.last_name = existing_tenant.last_name
                    End If
                
                    tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.serial + display_tenant.room_num
                    tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.first_name + display_tenant.middle_name + display_tenant.last_name
                    tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.address_district + display_tenant.address_municipality + display_tenant.address_ward
                    tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.contact_num + display_tenant.citizenship
                    tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.rent_year(0) + display_tenant.rent_month(0)
                                
                    If existing_tenant.rent_year(1) = 0 And existing_tenant.rent_month(1) = 0 Then 'tenant till living
                        tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + "Still living"
                    Else
                        tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.rent_year(1) + display_tenant.rent_month(1)
                    End If
                    tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + vbNewLine
                Next counter_tenant_1
            Close #1
        End If
    ElseIf tenant_filter_option(0).Value = Checked And tenant_filter_option(1).Value = Unchecked Then 'filter based on room status
        tenant_filter_status.Caption = "Filter status : On"

        tenant_fr_TextBox.Text = ""
        counter_tenant_2 = tenant_detail_count_function
        
        If counter_tenant_2 = 0 Then
            tenant_fr_TextBox.Text = "No detail has been added yet!"
        Else
            serial = 1
            Open "TenantDetail.txt" For Random As #1 Len = 103
                For counter_tenant_1 = 1 To counter_tenant_2 Step 1
                    Get #1, counter_tenant_1, existing_tenant
                    
                    display_tenant.serial = CStr(serial)
                    display_tenant.room_num = CStr(existing_tenant.room_num)
                    display_tenant.first_name = existing_tenant.first_name
                    display_tenant.address_district = existing_tenant.address_district
                    display_tenant.address_municipality = existing_tenant.address_municipality
                    display_tenant.address_ward = existing_tenant.address_ward
                    display_tenant.citizenship = existing_tenant.citizenship
                    display_tenant.contact_num = existing_tenant.contact_num
                    
                    display_tenant.rent_year(0) = existing_tenant.rent_year(0)
                    display_tenant.rent_year(1) = existing_tenant.rent_year(1)
                    
                    Call return_month_in_string(display_tenant.rent_month(0), existing_tenant.rent_month(0))
                    Call return_month_in_string(display_tenant.rent_month(1), existing_tenant.rent_month(1))
                    
                    If existing_tenant.middle_name = "Null      " Then ' has no middle name
                        display_tenant.middle_name = existing_tenant.last_name
                        display_tenant.last_name = ""
                    Else
                        display_tenant.middle_name = existing_tenant.middle_name
                        display_tenant.last_name = existing_tenant.last_name
                    End If
                    
                    ''''''filter based on room status
                    If tenant_filter_option_status(0).Value = True And tenant_filter_option_status(1).Value = False Then  ' still living
                        If existing_tenant.rent_year(1) = 0 And existing_tenant.rent_month(1) = 0 Then
                            tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.serial + display_tenant.room_num
                            tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.first_name + display_tenant.middle_name + display_tenant.last_name
                            tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.address_district + display_tenant.address_municipality + display_tenant.address_ward
                            tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.contact_num + display_tenant.citizenship
                            tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.rent_year(0) + display_tenant.rent_month(0) + "Still living" + vbNewLine
                            serial = serial + 1
                        End If
                    ElseIf tenant_filter_option_status(0).Value = False And tenant_filter_option_status(1).Value = True Then 'has already left
                        If existing_tenant.rent_year(1) <> 0 And existing_tenant.rent_month(1) <> 0 Then
                            tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.serial + display_tenant.room_num
                            tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.first_name + display_tenant.middle_name + display_tenant.last_name
                            tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.address_district + display_tenant.address_municipality + display_tenant.address_ward
                            tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.contact_num + display_tenant.citizenship
                            tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.rent_year(0) + display_tenant.rent_month(0) + display_tenant.rent_year(1) + display_tenant.rent_month(1) + vbNewLine
                            serial = serial + 1
                        End If
                    End If
                Next counter_tenant_1
            Close #1
        End If
    ElseIf tenant_filter_option(0).Value = Unchecked And tenant_filter_option(1).Value = Checked Then 'filter based on room number
        temp_integer = Val(tenant_filter_option_room.Text)
        tenant_filter_status.Caption = "Filter status : On"
        counter_tenant_2 = tenant_detail_count_function
        
        If counter_tenant_2 = 0 Then
            tenant_fr_TextBox.Text = "No detail has been added yet!"
        ElseIf counter_tenant_2 > 0 And tenant_filter_option_room.Text <> "" Then
            serial = 1
            tenant_fr_TextBox.Text = ""
            Open "TenantDetail.txt" For Random As #1 Len = 103
                For counter_tenant_1 = 1 To counter_tenant_2 Step 1
                    Get #1, counter_tenant_1, existing_tenant
                    If temp_integer = existing_tenant.room_num Then
                        display_tenant.serial = CStr(serial)
                        display_tenant.room_num = CStr(existing_tenant.room_num)
                        display_tenant.first_name = existing_tenant.first_name
                        display_tenant.address_district = existing_tenant.address_district
                        display_tenant.address_municipality = existing_tenant.address_municipality
                        display_tenant.address_ward = existing_tenant.address_ward
                        display_tenant.citizenship = existing_tenant.citizenship
                        display_tenant.contact_num = existing_tenant.contact_num
                        
                        display_tenant.rent_year(0) = existing_tenant.rent_year(0)
                        display_tenant.rent_year(1) = existing_tenant.rent_year(1)
                        
                        Call return_month_in_string(display_tenant.rent_month(0), existing_tenant.rent_month(0))
                        Call return_month_in_string(display_tenant.rent_month(1), existing_tenant.rent_month(1))
                
                        If existing_tenant.middle_name = "Null      " Then ' has no middle name
                            display_tenant.middle_name = existing_tenant.last_name
                            display_tenant.last_name = ""
                        Else
                            display_tenant.middle_name = existing_tenant.middle_name
                            display_tenant.last_name = existing_tenant.last_name
                        End If
                    
                        tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.serial + display_tenant.room_num
                        tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.first_name + display_tenant.middle_name + display_tenant.last_name
                        tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.address_district + display_tenant.address_municipality + display_tenant.address_ward
                        tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.contact_num + display_tenant.citizenship
                        tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.rent_year(0) + display_tenant.rent_month(0)
                                    
                        If existing_tenant.rent_year(1) = 0 And existing_tenant.rent_month(1) = 0 Then 'tenant till living
                            tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + "Still living"
                        Else
                            tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.rent_year(1) + display_tenant.rent_month(1)
                        End If
                        tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + vbNewLine
                        serial = serial + 1
                    End If
                Next counter_tenant_1
            Close #1
        Else
            MsgBox_Response = MsgBox("     Please select the room number first.", vbInformation + vbOKOnly, "Rental Record")
        End If
    ElseIf tenant_filter_option(0).Value = Checked And tenant_filter_option(1).Value = Checked Then 'filter based on room status and room number
        If tenant_filter_option_room.Text = "" Then
            MsgBox_Response = MsgBox("     Please enter the room number first.", vbInformation + vbOKOnly, "Rental Record")
            Exit Sub
        End If
        tenant_filter_status.Caption = "Filter status : On"
        tenant_fr_TextBox.Text = ""
        
        counter_tenant_2 = tenant_detail_count_function
        
        If counter_tenant_2 = 0 Then
            tenant_fr_TextBox.Text = "No detail has been added yet!"
        Else
            Open "TenantDetail.txt" For Random As #1 Len = 103
                serial = 1
                For counter_tenant_1 = 1 To counter_tenant_2 Step 1
                    Get #1, counter_tenant_1, existing_tenant
                    'selecting room number
                    If existing_tenant.room_num = Val(tenant_filter_option_room.Text) Then
                        'selecting room status
                        If tenant_filter_option_status(0).Value = True Then 'still living option selected
                            If existing_tenant.rent_year(1) = 0 And existing_tenant.rent_month(1) = 0 Then
                                display_tenant.serial = CStr(serial)
                                display_tenant.room_num = CStr(existing_tenant.room_num)
                                display_tenant.first_name = existing_tenant.first_name
                                display_tenant.address_district = existing_tenant.address_district
                                display_tenant.address_municipality = existing_tenant.address_municipality
                                display_tenant.address_ward = existing_tenant.address_ward
                                display_tenant.citizenship = existing_tenant.citizenship
                                display_tenant.contact_num = existing_tenant.contact_num
                                display_tenant.rent_year(0) = existing_tenant.rent_year(0)
                                display_tenant.rent_year(1) = existing_tenant.rent_year(1)
                                
                               
                                Call return_month_in_string(display_tenant.rent_month(0), existing_tenant.rent_month(0))
                                Call return_month_in_string(display_tenant.rent_month(1), existing_tenant.rent_month(1))
                            
                                If existing_tenant.middle_name = "Null      " Then ' has no middle name
                                    display_tenant.middle_name = existing_tenant.last_name
                                    display_tenant.last_name = ""
                                Else
                                    display_tenant.middle_name = existing_tenant.middle_name
                                    display_tenant.last_name = existing_tenant.last_name
                                End If
                                
                                tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.serial + display_tenant.room_num
                                tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.first_name + display_tenant.middle_name + display_tenant.last_name
                                tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.address_district + display_tenant.address_municipality + display_tenant.address_ward
                                tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.contact_num + display_tenant.citizenship
                                tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.rent_year(0) + display_tenant.rent_month(0) + "Still living" + vbNewLine
                                
                                serial = serial + 1
                            End If
                        Else
                            If existing_tenant.rent_year(1) <> 0 And existing_tenant.rent_month(1) <> 0 Then
                                display_tenant.serial = CStr(serial)
                                display_tenant.room_num = CStr(existing_tenant.room_num)
                                display_tenant.first_name = existing_tenant.first_name
                                display_tenant.address_district = existing_tenant.address_district
                                display_tenant.address_municipality = existing_tenant.address_municipality
                                display_tenant.address_ward = existing_tenant.address_ward
                                display_tenant.citizenship = existing_tenant.citizenship
                                display_tenant.contact_num = existing_tenant.contact_num
                                display_tenant.rent_year(0) = existing_tenant.rent_year(0)
                                display_tenant.rent_year(1) = existing_tenant.rent_year(1)
                                
                                Call return_month_in_string(display_tenant.rent_month(0), existing_tenant.rent_month(0))
                                Call return_month_in_string(display_tenant.rent_month(1), existing_tenant.rent_month(1))
                                
                                If existing_tenant.middle_name = "Null      " Then ' has no middle name
                                    display_tenant.middle_name = existing_tenant.last_name
                                    display_tenant.last_name = ""
                                Else
                                    display_tenant.middle_name = existing_tenant.middle_name
                                    display_tenant.last_name = existing_tenant.last_name
                                End If
                                
                                tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.serial + display_tenant.room_num
                                tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.first_name + display_tenant.middle_name + display_tenant.last_name
                                tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.address_district + display_tenant.address_municipality + display_tenant.address_ward
                                tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.contact_num + display_tenant.citizenship
                                tenant_fr_TextBox.Text = tenant_fr_TextBox.Text + display_tenant.rent_year(0) + display_tenant.rent_month(0) + display_tenant.rent_year(1) + display_tenant.rent_month(1) + vbNewLine
                                serial = serial + 1
                            End If
                        End If
                    End If
                Next counter_tenant_1
            Close #1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Main_Form
End Sub

