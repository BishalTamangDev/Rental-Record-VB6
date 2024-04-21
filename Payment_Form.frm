VERSION 5.00
Begin VB.Form Payment_Form 
   Caption         =   "Rental Record"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   LinkTopic       =   "Form1"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame payment_frame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Payment"
      Height          =   12375
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   23055
      Begin VB.Frame payment_selection_fr 
         BackColor       =   &H8000000E&
         Height          =   5175
         Left            =   9180
         TabIndex        =   18
         Top             =   5520
         Width           =   4455
         Begin VB.TextBox payment_unit_TextBox 
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
            Left            =   1800
            TabIndex        =   23
            Text            =   "Enter electricity unit"
            Top             =   2760
            Width           =   2415
         End
         Begin VB.CommandButton payment_next_command 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            TabIndex        =   20
            Top             =   3600
            Width           =   3975
         End
         Begin VB.ComboBox payment_room_combo 
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
            ItemData        =   "Payment_Form.frx":0000
            Left            =   1800
            List            =   "Payment_Form.frx":000D
            Sorted          =   -1  'True
            TabIndex        =   19
            Text            =   "Room number"
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter the electricit unit consumed and click on proceed."
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Index           =   30
            Left            =   240
            TabIndex        =   209
            Top             =   2040
            Width           =   4095
            WordWrap        =   -1  'True
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select the room number and press check for checking if the rent is pending.  "
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Index           =   9
            Left            =   240
            TabIndex        =   208
            Top             =   360
            Width           =   4545
            WordWrap        =   -1  'True
         End
         Begin VB.Label payment_goback_1 
            AutoSize        =   -1  'True
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
            Height          =   315
            Left            =   1800
            TabIndex        =   136
            Top             =   4440
            Width           =   885
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Room number"
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
            Left            =   240
            TabIndex        =   133
            Top             =   1200
            Width           =   1260
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Electricity Unit"
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
            Left            =   240
            TabIndex        =   108
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Image Image3 
            Height          =   615
            Left            =   240
            Picture         =   "Payment_Form.frx":001A
            Stretch         =   -1  'True
            Top             =   4320
            Width           =   3975
         End
      End
      Begin VB.Frame payment_confirm_fr 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Caption         =   "Payment confirmation frame"
         Height          =   7935
         Left            =   4793
         TabIndex        =   21
         Top             =   3600
         Width           =   13215
         Begin VB.CheckBox payment_service_check 
            BackColor       =   &H8000000E&
            Caption         =   "Waste Management"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   3000
            TabIndex        =   141
            Top             =   5160
            Width           =   2415
         End
         Begin VB.CheckBox payment_service_check 
            BackColor       =   &H8000000E&
            Caption         =   "Drinking water"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   3000
            TabIndex        =   140
            Top             =   4680
            Width           =   2055
         End
         Begin VB.CheckBox payment_service_check 
            BackColor       =   &H8000000E&
            Caption         =   "Electricity"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   3000
            TabIndex        =   139
            Top             =   4200
            Width           =   1455
         End
         Begin VB.CheckBox payment_service_check 
            BackColor       =   &H8000000E&
            Caption         =   "Security"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   3000
            TabIndex        =   138
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CheckBox payment_service_check 
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
            Height          =   315
            Index           =   0
            Left            =   3000
            TabIndex        =   137
            Top             =   3240
            Width           =   1215
         End
         Begin VB.CommandButton payment_pay 
            Caption         =   "Pay"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   7200
            TabIndex        =   22
            Top             =   6960
            Width           =   5175
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "xxxx"
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
            Index           =   43
            Left            =   3000
            TabIndex        =   222
            Top             =   2040
            Width           =   480
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
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
            Index           =   42
            Left            =   9840
            TabIndex        =   221
            Top             =   1080
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rs."
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
            Index           =   41
            Left            =   9840
            TabIndex        =   220
            Top             =   6360
            Width           =   255
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rs."
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
            Index           =   40
            Left            =   9840
            TabIndex        =   219
            Top             =   5640
            Width           =   255
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rs."
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
            Index           =   39
            Left            =   9840
            TabIndex        =   218
            Top             =   4440
            Width           =   255
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rs."
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
            Index           =   38
            Left            =   9840
            TabIndex        =   217
            Top             =   3840
            Width           =   255
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rs."
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
            Index           =   37
            Left            =   9840
            TabIndex        =   216
            Top             =   3240
            Width           =   255
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rs."
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
            Index           =   36
            Left            =   9840
            TabIndex        =   215
            Top             =   2640
            Width           =   255
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rs."
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
            Index           =   35
            Left            =   9840
            TabIndex        =   214
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
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
            Index           =   34
            Left            =   5280
            TabIndex        =   213
            Top             =   1080
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "xxxx"
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
            Index           =   33
            Left            =   4080
            TabIndex        =   212
            Top             =   1080
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
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
            Index           =   32
            Left            =   7200
            TabIndex        =   211
            Top             =   6360
            Width           =   435
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
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
            Index           =   31
            Left            =   10200
            TabIndex        =   210
            Top             =   6360
            Width           =   450
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000000&
            Height          =   7335
            Left            =   480
            Top             =   480
            Width           =   5895
         End
         Begin VB.Label payment_goback_2 
            AutoSize        =   -1  'True
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
            Height          =   315
            Left            =   3000
            TabIndex        =   135
            Top             =   7080
            Width           =   885
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
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
            Index           =   29
            Left            =   10200
            TabIndex        =   134
            Top             =   5640
            Width           =   450
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "xxx"
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
            Index           =   28
            Left            =   9840
            TabIndex        =   132
            Top             =   5040
            Width           =   270
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
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
            Index           =   27
            Left            =   10200
            TabIndex        =   131
            Top             =   4440
            Width           =   450
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
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
            Index           =   26
            Left            =   10200
            TabIndex        =   130
            Top             =   3840
            Width           =   450
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
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
            Index           =   25
            Left            =   10200
            TabIndex        =   129
            Top             =   3240
            Width           =   450
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
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
            Index           =   24
            Left            =   10200
            TabIndex        =   128
            Top             =   2640
            Width           =   450
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
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
            Index           =   23
            Left            =   10200
            TabIndex        =   127
            Top             =   2040
            Width           =   450
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "xxx xxxx"
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
            Index           =   22
            Left            =   9840
            TabIndex        =   126
            Top             =   1440
            Width           =   690
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Electricity fee"
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
            Index           =   21
            Left            =   7200
            TabIndex        =   125
            Top             =   5640
            Width           =   1140
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Waste management"
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
            Index           =   20
            Left            =   7200
            TabIndex        =   124
            Top             =   3240
            Width           =   1740
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Index           =   19
            Left            =   7200
            TabIndex        =   123
            Top             =   4440
            Width           =   660
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Drinking water"
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
            Left            =   7200
            TabIndex        =   122
            Top             =   2640
            Width           =   1260
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Electricity unit"
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
            Left            =   7200
            TabIndex        =   121
            Top             =   5040
            Width           =   1185
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Security"
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
            Left            =   7200
            TabIndex        =   120
            Top             =   3840
            Width           =   675
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Amount"
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
            Left            =   7200
            TabIndex        =   119
            Top             =   2040
            Width           =   1125
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rent for"
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
            Left            =   7200
            TabIndex        =   118
            Top             =   1440
            Width           =   705
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bill"
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
            Index           =   13
            Left            =   9600
            TabIndex        =   117
            Top             =   600
            Width           =   330
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
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
            Index           =   12
            Left            =   3000
            TabIndex        =   116
            Top             =   2640
            Width           =   360
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "xxxx"
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
            Index           =   11
            Left            =   3000
            TabIndex        =   115
            Top             =   1080
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
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
            Left            =   3000
            TabIndex        =   114
            Top             =   1440
            Width           =   360
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service provided"
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
            Left            =   840
            TabIndex        =   113
            Top             =   3240
            Width           =   1470
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
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
            Index           =   7
            Left            =   840
            TabIndex        =   112
            Top             =   2640
            Width           =   1395
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tenant name"
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
            Left            =   840
            TabIndex        =   111
            Top             =   2040
            Width           =   1125
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Room number"
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
            Left            =   840
            TabIndex        =   110
            Top             =   1440
            Width           =   1260
         End
         Begin VB.Label payment_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tenant Details"
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
            Index           =   4
            Left            =   2640
            TabIndex        =   109
            Top             =   600
            Width           =   1545
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00800000&
            X1              =   7200
            X2              =   12360
            Y1              =   6120
            Y2              =   6120
         End
         Begin VB.Image Image4 
            Appearance      =   0  'Flat
            Height          =   615
            Left            =   6840
            Picture         =   "Payment_Form.frx":0525
            Stretch         =   -1  'True
            Top             =   480
            Width           =   5895
         End
         Begin VB.Image Image2 
            Appearance      =   0  'Flat
            Height          =   615
            Left            =   480
            Picture         =   "Payment_Form.frx":0A2F
            Stretch         =   -1  'True
            Top             =   480
            Width           =   5895
         End
         Begin VB.Image Image5 
            Height          =   615
            Left            =   720
            Picture         =   "Payment_Form.frx":0F39
            Stretch         =   -1  'True
            Top             =   6960
            Width           =   5415
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H80000000&
            Height          =   7335
            Left            =   6840
            Top             =   480
            Width           =   5895
         End
      End
      Begin VB.Label payment_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Do Payment"
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
         Left            =   10395
         TabIndex        =   107
         Top             =   2760
         Width           =   2010
      End
      Begin VB.Label payment_fr_label 
         AutoSize        =   -1  'True
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
         TabIndex        =   106
         Top             =   480
         Width           =   4695
      End
      Begin VB.Image Image22 
         Height          =   1050
         Left            =   5520
         Picture         =   "Payment_Form.frx":1444
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Image Image14 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   720
         Picture         =   "Payment_Form.frx":B14E
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   21615
      End
      Begin VB.Image Image6 
         Height          =   615
         Index           =   0
         Left            =   6000
         Picture         =   "Payment_Form.frx":B658
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   735
      End
   End
   Begin VB.Frame service_add_fr 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Service charge addition"
      Height          =   12375
      Left            =   0
      TabIndex        =   62
      Top             =   0
      Width           =   23055
      Begin VB.ComboBox fee_combo 
         BackColor       =   &H80000014&
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
         ItemData        =   "Payment_Form.frx":11DA4
         Left            =   10320
         List            =   "Payment_Form.frx":11DF0
         TabIndex        =   151
         Text            =   "Year"
         Top             =   8160
         Width           =   1335
      End
      Begin VB.TextBox fee_textbox 
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
         Left            =   12000
         TabIndex        =   150
         Text            =   "Enter fee"
         Top             =   5160
         Width           =   2535
      End
      Begin VB.TextBox fee_textbox 
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
         Left            =   12000
         TabIndex        =   149
         Text            =   "Enter fee"
         Top             =   5880
         Width           =   2535
      End
      Begin VB.TextBox fee_textbox 
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
         Left            =   12000
         TabIndex        =   148
         Text            =   "Enter fee"
         Top             =   6600
         Width           =   2535
      End
      Begin VB.TextBox fee_textbox 
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
         Left            =   12000
         TabIndex        =   147
         Text            =   "Enter fee"
         Top             =   7320
         Width           =   2535
      End
      Begin VB.ComboBox fee_combo 
         BackColor       =   &H80000014&
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
         ItemData        =   "Payment_Form.frx":11E84
         Left            =   12000
         List            =   "Payment_Form.frx":11EAF
         TabIndex        =   146
         Text            =   "Month"
         Top             =   8160
         Width           =   2550
      End
      Begin VB.CommandButton service_add_fr_save 
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
         Height          =   615
         Left            =   8273
         TabIndex        =   142
         Top             =   8880
         Width           =   6255
      End
      Begin VB.Label service_add_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prevailing fee"
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
         Left            =   10320
         TabIndex        =   145
         Top             =   4320
         Width           =   1485
      End
      Begin VB.Label service_add_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service"
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
         Left            =   8280
         TabIndex        =   144
         Top             =   4320
         Width           =   795
      End
      Begin VB.Label service_add_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Fee"
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
         Index           =   4
         Left            =   12000
         TabIndex        =   143
         Top             =   4320
         Width           =   945
      End
      Begin VB.Image Image19 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   8040
         Picture         =   "Payment_Form.frx":11F17
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   6615
      End
      Begin VB.Label service_add_fr_goback 
         AutoSize        =   -1  'True
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
         Height          =   315
         Left            =   10920
         TabIndex        =   161
         Top             =   9840
         Width           =   885
      End
      Begin VB.Label service_add_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drinking water"
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
         Left            =   8280
         TabIndex        =   160
         Top             =   5280
         Width           =   1260
      End
      Begin VB.Label service_add_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Waste management"
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
         Left            =   8280
         TabIndex        =   159
         Top             =   6000
         Width           =   1740
      End
      Begin VB.Label service_add_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Security"
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
         Left            =   8280
         TabIndex        =   158
         Top             =   6720
         Width           =   675
      End
      Begin VB.Label service_add_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Index           =   8
         Left            =   8280
         TabIndex        =   157
         Top             =   7440
         Width           =   660
      End
      Begin VB.Label service_add_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set a effective date : "
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   13
         Left            =   8280
         TabIndex        =   156
         Top             =   8160
         Width           =   1890
      End
      Begin VB.Label service_add_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs. xxx"
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
         Left            =   10320
         TabIndex        =   155
         Top             =   5280
         Width           =   615
      End
      Begin VB.Label service_add_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs. xxx"
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
         Left            =   10320
         TabIndex        =   154
         Top             =   6000
         Width           =   585
      End
      Begin VB.Label service_add_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs. xxx"
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
         Left            =   10320
         TabIndex        =   153
         Top             =   6720
         Width           =   585
      End
      Begin VB.Label service_add_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs. xxx"
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
         Left            =   10320
         TabIndex        =   152
         Top             =   7440
         Width           =   585
      End
      Begin VB.Image Image18 
         Height          =   615
         Left            =   8280
         Picture         =   "Payment_Form.frx":12421
         Stretch         =   -1  'True
         Top             =   9720
         Width           =   6255
      End
      Begin VB.Label service_add_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add New Service Fee"
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
         Left            =   9893
         TabIndex        =   65
         Top             =   2520
         Width           =   3015
      End
      Begin VB.Label service_add_fr_label 
         AutoSize        =   -1  'True
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
         TabIndex        =   64
         Top             =   480
         Width           =   4605
      End
      Begin VB.Label issue_detail_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
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
         Index           =   44
         Left            =   -360
         TabIndex        =   63
         Top             =   360
         Width           =   1560
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image20 
         Height          =   1050
         Left            =   5520
         Picture         =   "Payment_Form.frx":1292C
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Image Image21 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   720
         Picture         =   "Payment_Form.frx":1C636
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   21495
      End
      Begin VB.Image Image23 
         Height          =   8655
         Left            =   6000
         Picture         =   "Payment_Form.frx":1CB40
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   10695
      End
   End
   Begin VB.Frame electricity_fee_frame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Service charge detail"
      Height          =   12375
      Left            =   0
      TabIndex        =   60
      Top             =   0
      Width           =   23055
      Begin VB.Frame Frame38 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   7920
         TabIndex        =   202
         Top             =   8760
         Width           =   6975
      End
      Begin VB.Frame Frame37 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   7920
         TabIndex        =   201
         Top             =   9480
         Width           =   6975
      End
      Begin VB.Frame Frame52 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   4695
         Left            =   13200
         TabIndex        =   200
         Top             =   4080
         Width           =   15
      End
      Begin VB.Frame Frame55 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   5415
         Left            =   11400
         TabIndex        =   199
         Top             =   4080
         Width           =   15
      End
      Begin VB.Frame Frame47 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   7800
         TabIndex        =   198
         Top             =   4560
         Width           =   3615
      End
      Begin VB.Frame Frame54 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   9600
         TabIndex        =   197
         Top             =   4560
         Width           =   15
      End
      Begin VB.CommandButton electricity_fee_save 
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
         Height          =   615
         Left            =   7920
         TabIndex        =   188
         Top             =   9720
         Width           =   6975
      End
      Begin VB.TextBox unit 
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
         Index           =   23
         Left            =   13320
         TabIndex        =   187
         Text            =   "per unit"
         Top             =   8280
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Index           =   22
         Left            =   11520
         TabIndex        =   186
         Text            =   "monthly min"
         Top             =   8280
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Index           =   21
         Left            =   9720
         TabIndex        =   185
         Text            =   "max"
         Top             =   8280
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Index           =   20
         Left            =   7920
         TabIndex        =   184
         Text            =   "min"
         Top             =   8280
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Index           =   19
         Left            =   13320
         TabIndex        =   183
         Text            =   "per unit"
         Top             =   7680
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Index           =   18
         Left            =   11520
         TabIndex        =   182
         Text            =   "monthly min"
         Top             =   7680
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Index           =   17
         Left            =   9720
         TabIndex        =   181
         Text            =   "max"
         Top             =   7680
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Index           =   16
         Left            =   7920
         TabIndex        =   180
         Text            =   "min"
         Top             =   7680
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Index           =   15
         Left            =   13320
         TabIndex        =   179
         Text            =   "per unit"
         Top             =   7080
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Index           =   14
         Left            =   11520
         TabIndex        =   178
         Text            =   "monthly min"
         Top             =   7080
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Index           =   13
         Left            =   9720
         TabIndex        =   177
         Text            =   "max"
         Top             =   7080
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Index           =   12
         Left            =   7920
         TabIndex        =   176
         Text            =   "min"
         Top             =   7080
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Index           =   11
         Left            =   13320
         TabIndex        =   175
         Text            =   "per unit"
         Top             =   6480
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Index           =   10
         Left            =   11520
         TabIndex        =   174
         Text            =   "monthly min"
         Top             =   6480
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Index           =   9
         Left            =   9720
         TabIndex        =   173
         Text            =   "max"
         Top             =   6480
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Index           =   8
         Left            =   7920
         TabIndex        =   172
         Text            =   "min"
         Top             =   6480
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Index           =   7
         Left            =   13320
         TabIndex        =   171
         Text            =   "per unit"
         Top             =   5880
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Index           =   6
         Left            =   11520
         TabIndex        =   170
         Text            =   "monthly min"
         Top             =   5880
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Index           =   5
         Left            =   9720
         TabIndex        =   169
         Text            =   "max"
         Top             =   5880
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Left            =   7920
         TabIndex        =   168
         Text            =   "min"
         Top             =   5880
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Left            =   13320
         TabIndex        =   167
         Text            =   "per unit"
         Top             =   5280
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Left            =   11520
         TabIndex        =   166
         Text            =   "monthly min"
         Top             =   5280
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Left            =   9720
         TabIndex        =   165
         Text            =   "max"
         Top             =   5280
         Width           =   1530
      End
      Begin VB.TextBox unit 
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
         Left            =   7920
         TabIndex        =   164
         Text            =   "min"
         Top             =   5280
         Width           =   1530
      End
      Begin VB.ComboBox unit_combo 
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
         ItemData        =   "Payment_Form.frx":2328C
         Left            =   11520
         List            =   "Payment_Form.frx":232D8
         TabIndex        =   163
         Text            =   "Year"
         Top             =   8880
         Width           =   1575
      End
      Begin VB.ComboBox unit_combo 
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
         ItemData        =   "Payment_Form.frx":2336C
         Left            =   13320
         List            =   "Payment_Form.frx":23394
         TabIndex        =   162
         Text            =   "Month"
         Top             =   8880
         Width           =   1575
      End
      Begin VB.Label electricity_fee_goback 
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
         Left            =   10920
         TabIndex        =   196
         Top             =   10560
         Width           =   885
      End
      Begin VB.Label electricity_fee_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Energy fee (per kilo watt hour)"
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
         Index           =   6
         Left            =   13320
         TabIndex        =   195
         Top             =   4200
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label electricity_fee_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly minimum fee"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   5
         Left            =   11520
         TabIndex        =   194
         Top             =   4320
         Width           =   1530
         WordWrap        =   -1  'True
      End
      Begin VB.Label electricity_fee_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum"
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
         Left            =   9720
         TabIndex        =   193
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label electricity_fee_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum"
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
         Left            =   7920
         TabIndex        =   192
         Top             =   4680
         Width           =   810
      End
      Begin VB.Label electricity_fee_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kilowatt hour unit (Range)"
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
         Left            =   7920
         TabIndex        =   191
         Top             =   4200
         Width           =   2265
      End
      Begin VB.Label electricity_fee_extract 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Extract latest fee detail"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   190
         Top             =   11160
         Width           =   2070
      End
      Begin VB.Label electricity_fee_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select the effective date :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   7920
         TabIndex        =   189
         Top             =   9000
         Width           =   2550
      End
      Begin VB.Label electricity_fee_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add New Electricity Unit"
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
         Height          =   495
         Index           =   1
         Left            =   9420
         TabIndex        =   67
         Top             =   2520
         Width           =   3975
      End
      Begin VB.Label electricity_fee_label 
         AutoSize        =   -1  'True
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
         TabIndex        =   66
         Top             =   480
         Width           =   4695
      End
      Begin VB.Image Image15 
         Height          =   1050
         Left            =   5520
         Picture         =   "Payment_Form.frx":233FA
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label issue_detail_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
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
         Index           =   46
         Left            =   -360
         TabIndex        =   61
         Top             =   360
         Width           =   1560
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image16 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   720
         Picture         =   "Payment_Form.frx":2D104
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   21495
      End
      Begin VB.Image Image13 
         Height          =   615
         Left            =   7920
         Picture         =   "Payment_Form.frx":2D60E
         Stretch         =   -1  'True
         Top             =   10440
         Width           =   6975
      End
      Begin VB.Image Image12 
         Height          =   975
         Left            =   7680
         Picture         =   "Payment_Form.frx":2DB19
         Stretch         =   -1  'True
         Top             =   4080
         Width           =   7455
      End
      Begin VB.Image Image24 
         Height          =   11175
         Left            =   5400
         Picture         =   "Payment_Form.frx":2E024
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   12015
      End
   End
   Begin VB.Frame service_det_frame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Service charge detail"
      Height          =   12375
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   23055
      Begin VB.Frame Frame21 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Caption         =   "Service charge detail frame"
         Height          =   7335
         Left            =   600
         TabIndex        =   34
         Top             =   3240
         Width           =   11295
         Begin VB.Frame Frame34 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   7095
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   15
         End
         Begin VB.Frame Frame26 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            ForeColor       =   &H8000000A&
            Height          =   7095
            Left            =   840
            TabIndex        =   44
            Top             =   240
            Width           =   15
         End
         Begin VB.Frame Frame29 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   7095
            Left            =   3120
            TabIndex        =   43
            Top             =   240
            Width           =   15
         End
         Begin VB.Frame Frame22 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   7215
            Left            =   5400
            TabIndex        =   42
            Top             =   240
            Width           =   15
         End
         Begin VB.Frame Frame23 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   7095
            Left            =   8280
            TabIndex        =   41
            Top             =   240
            Width           =   15
         End
         Begin VB.Frame Frame24 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   7095
            Left            =   6840
            TabIndex        =   40
            Top             =   240
            Width           =   15
         End
         Begin VB.Frame Frame27 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   7095
            Left            =   9720
            TabIndex        =   39
            Top             =   240
            Width           =   15
         End
         Begin VB.Frame Frame33 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   15
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   11055
         End
         Begin VB.Frame Frame28 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   15
            Left            =   120
            TabIndex        =   37
            Top             =   1200
            Width           =   11055
         End
         Begin VB.Frame Frame39 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   15
            Left            =   120
            TabIndex        =   36
            Top             =   7320
            Width           =   11055
         End
         Begin VB.Frame Frame40 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   7095
            Left            =   11160
            TabIndex        =   35
            Top             =   240
            Width           =   15
         End
         Begin VB.TextBox Service_fee_TextBox 
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
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   46
            Text            =   "Payment_Form.frx":34770
            Top             =   1320
            Width           =   10935
         End
         Begin VB.Label service_det_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Index           =   7
            Left            =   9840
            TabIndex        =   82
            Top             =   360
            Width           =   660
         End
         Begin VB.Label service_det_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Security"
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
            Left            =   8400
            TabIndex        =   81
            Top             =   360
            Width           =   675
         End
         Begin VB.Label service_det_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Waste Management"
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
            Index           =   5
            Left            =   6960
            TabIndex        =   80
            Top             =   360
            Width           =   1155
            WordWrap        =   -1  'True
         End
         Begin VB.Label service_det_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Drinking Water"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   4
            Left            =   5520
            TabIndex        =   79
            Top             =   360
            Width           =   825
            WordWrap        =   -1  'True
         End
         Begin VB.Label service_det_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Date"
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
            Left            =   3240
            TabIndex        =   78
            Top             =   360
            Width           =   795
         End
         Begin VB.Label service_det_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date"
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
            Left            =   960
            TabIndex        =   77
            Top             =   360
            Width           =   870
         End
         Begin VB.Label service_det_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S.N."
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
            Left            =   240
            TabIndex        =   76
            Top             =   360
            Width           =   345
         End
         Begin VB.Image Image9 
            Height          =   975
            Left            =   120
            Picture         =   "Payment_Form.frx":347C7
            Stretch         =   -1  'True
            Top             =   240
            Width           =   11055
         End
      End
      Begin VB.Frame Frame20 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Caption         =   "Electricity charge detail"
         Height          =   7335
         Left            =   11880
         TabIndex        =   47
         Top             =   3240
         Width           =   10575
         Begin VB.Frame Frame19 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   7095
            Left            =   840
            TabIndex        =   86
            Top             =   240
            Width           =   15
         End
         Begin VB.Frame Frame30 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   15
            Left            =   5400
            TabIndex        =   59
            Top             =   840
            Width           =   1935
         End
         Begin VB.Frame Frame45 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   7095
            Left            =   10440
            TabIndex        =   57
            Top             =   240
            Width           =   15
         End
         Begin VB.Frame Frame44 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   15
            Left            =   120
            TabIndex        =   56
            Top             =   7320
            Width           =   10335
         End
         Begin VB.Frame Frame43 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   15
            Left            =   120
            TabIndex        =   55
            Top             =   1200
            Width           =   10335
         End
         Begin VB.Frame Frame42 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   15
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   10335
         End
         Begin VB.Frame Frame41 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   7095
            Left            =   8520
            TabIndex        =   53
            Top             =   240
            Width           =   15
         End
         Begin VB.Frame Frame36 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   7095
            Left            =   7320
            TabIndex        =   52
            Top             =   240
            Width           =   15
         End
         Begin VB.Frame Frame35 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   6495
            Left            =   6360
            TabIndex        =   51
            Top             =   840
            Width           =   15
         End
         Begin VB.Frame Frame32 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   7095
            Left            =   5400
            TabIndex        =   50
            Top             =   240
            Width           =   15
         End
         Begin VB.Frame Frame31 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   7095
            Left            =   3120
            TabIndex        =   49
            Top             =   240
            Width           =   15
         End
         Begin VB.Frame Frame25 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   7095
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   15
         End
         Begin VB.TextBox Electricity_fee_TextBox 
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
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   58
            Text            =   "Payment_Form.frx":34CD2
            Top             =   1320
            Width           =   10215
         End
         Begin VB.Label service_det_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S.N."
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
            Left            =   240
            TabIndex        =   85
            Top             =   360
            Width           =   345
         End
         Begin VB.Label service_det_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End date"
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
            Left            =   3240
            TabIndex        =   75
            Top             =   360
            Width           =   780
         End
         Begin VB.Label service_det_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start date"
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
            Left            =   960
            TabIndex        =   74
            Top             =   360
            Width           =   855
         End
         Begin VB.Label service_det_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
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
            Left            =   6480
            TabIndex        =   73
            Top             =   840
            Width           =   375
         End
         Begin VB.Label service_det_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
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
            Left            =   5520
            TabIndex        =   72
            Top             =   840
            Width           =   330
         End
         Begin VB.Label service_det_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kilo Watt hour unit"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Index           =   12
            Left            =   5520
            TabIndex        =   71
            Top             =   360
            Width           =   1680
            WordWrap        =   -1  'True
         End
         Begin VB.Label service_det_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monthly minimum fee"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Index           =   13
            Left            =   7440
            TabIndex        =   70
            Top             =   360
            Width           =   1050
            WordWrap        =   -1  'True
         End
         Begin VB.Label service_det_fr_label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Energy fee (per kilo watt hour unit)"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Index           =   14
            Left            =   8640
            TabIndex        =   69
            Top             =   360
            Width           =   1650
            WordWrap        =   -1  'True
         End
         Begin VB.Image Image11 
            Height          =   975
            Left            =   120
            Picture         =   "Payment_Form.frx":34D57
            Stretch         =   -1  'True
            Top             =   240
            Width           =   10335
         End
      End
      Begin VB.Label service_detail_fr_goback 
         AutoSize        =   -1  'True
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
         Height          =   375
         Left            =   1560
         TabIndex        =   84
         Top             =   11160
         Width           =   1035
      End
      Begin VB.Label service_det_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service Charge Details"
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
         Index           =   15
         Left            =   960
         TabIndex        =   83
         Top             =   2520
         Width           =   3645
      End
      Begin VB.Label service_det_fr_label 
         AutoSize        =   -1  'True
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
         TabIndex        =   68
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label issue_detail_label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
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
         Index           =   27
         Left            =   -360
         TabIndex        =   33
         Top             =   360
         Width           =   1560
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image7 
         Height          =   1050
         Left            =   5520
         Picture         =   "Payment_Form.frx":35262
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Image Image8 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   720
         Picture         =   "Payment_Form.frx":3EF6C
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   4095
      End
      Begin VB.Image Image10 
         Height          =   615
         Left            =   720
         Picture         =   "Payment_Form.frx":3F476
         Stretch         =   -1  'True
         Top             =   11040
         Width           =   2775
      End
   End
   Begin VB.Frame payment_detail_frame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Payment Detail Frame"
      Height          =   12375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   23055
      Begin VB.Frame Frame46 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   14040
         TabIndex        =   204
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   18240
         TabIndex        =   28
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   21000
         TabIndex        =   27
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   19800
         TabIndex        =   26
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   16680
         TabIndex        =   25
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   15480
         TabIndex        =   24
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   12840
         TabIndex        =   15
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   1440
         TabIndex        =   14
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   2160
         TabIndex        =   12
         Top             =   3480
         Width           =   15
         Begin VB.Frame Frame4 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   6735
            Left            =   120
            TabIndex        =   13
            Top             =   2520
            Width           =   135
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   5760
         TabIndex        =   11
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   7680
         TabIndex        =   10
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   9480
         TabIndex        =   9
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   11760
         TabIndex        =   8
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   720
         TabIndex        =   7
         Top             =   3480
         Width           =   21615
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   720
         TabIndex        =   6
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   720
         TabIndex        =   5
         Top             =   9840
         Width           =   21615
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   720
         TabIndex        =   4
         Top             =   4200
         Width           =   21615
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   22320
         TabIndex        =   3
         Top             =   3480
         Width           =   15
      End
      Begin VB.Frame payment_filter_frame 
         BorderStyle     =   0  'None
         Caption         =   "Tenant Filter Frame"
         Height          =   1575
         Left            =   15480
         TabIndex        =   1
         Top             =   10080
         Width           =   6855
         Begin VB.OptionButton payment_filter_option 
            Caption         =   "Unpaid"
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
            Left            =   5280
            TabIndex        =   207
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton payment_filter_option 
            Caption         =   "Paid"
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
            Left            =   5280
            TabIndex        =   206
            Top             =   120
            Width           =   735
         End
         Begin VB.CheckBox payment_filter_check 
            Caption         =   "Rent Status"
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
            Left            =   3600
            TabIndex        =   205
            Top             =   120
            Width           =   1335
         End
         Begin VB.ComboBox payment_filter_combo 
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
            ItemData        =   "Payment_Form.frx":3F981
            Left            =   1800
            List            =   "Payment_Form.frx":3F9A9
            TabIndex        =   105
            Text            =   "Month"
            Top             =   1080
            Width           =   1575
         End
         Begin VB.ComboBox payment_filter_combo 
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
            Left            =   1800
            TabIndex        =   104
            Text            =   "Year"
            Top             =   600
            Width           =   1575
         End
         Begin VB.ComboBox payment_filter_combo 
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
            Left            =   1800
            TabIndex        =   103
            Text            =   "Room No"
            Top             =   120
            Width           =   1575
         End
         Begin VB.CheckBox payment_filter_check 
            Caption         =   "Month"
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
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox payment_filter_check 
            Caption         =   "Year"
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
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox payment_filter_check 
            Caption         =   "Room number"
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
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   1575
         End
         Begin VB.CommandButton payment_filter_command 
            Caption         =   "Filter : Off"
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
            Left            =   3600
            TabIndex        =   2
            Top             =   960
            Width           =   3135
         End
      End
      Begin VB.TextBox payment_detail_TextBox 
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
         Height          =   5535
         Left            =   840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Text            =   "Payment_Form.frx":3FA0F
         Top             =   4440
         Width           =   21495
      End
      Begin VB.Label payment_detail_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment of "
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
         Index           =   15
         Left            =   7800
         TabIndex        =   203
         Top             =   3600
         Width           =   1440
         WordWrap        =   -1  'True
      End
      Begin VB.Label payment_detail_fr_goback 
         AutoSize        =   -1  'True
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
         Height          =   375
         Left            =   1560
         TabIndex        =   102
         Top             =   11160
         Width           =   1035
      End
      Begin VB.Label payment_detail_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Index           =   14
         Left            =   21000
         TabIndex        =   101
         Top             =   3600
         Width           =   600
      End
      Begin VB.Label payment_detail_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Internet"
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
         Index           =   13
         Left            =   19920
         TabIndex        =   100
         Top             =   3600
         Width           =   960
      End
      Begin VB.Label payment_detail_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Electricity Charge"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   12
         Left            =   18360
         TabIndex        =   99
         Top             =   3600
         Width           =   1440
         WordWrap        =   -1  'True
      End
      Begin VB.Label payment_detail_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Electricity Unit"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   11
         Left            =   16800
         TabIndex        =   98
         Top             =   3600
         Width           =   1440
         WordWrap        =   -1  'True
      End
      Begin VB.Label payment_detail_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Security"
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
         Index           =   10
         Left            =   15600
         TabIndex        =   97
         Top             =   3600
         Width           =   960
      End
      Begin VB.Label payment_detail_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Waste Management"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   9
         Left            =   14160
         TabIndex        =   96
         Top             =   3600
         Width           =   1680
         WordWrap        =   -1  'True
      End
      Begin VB.Label payment_detail_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drinking Water"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   8
         Left            =   12960
         TabIndex        =   95
         Top             =   3600
         Width           =   1440
         WordWrap        =   -1  'True
      End
      Begin VB.Label payment_detail_fr_label 
         AutoSize        =   -1  'True
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
         Height          =   480
         Index           =   7
         Left            =   11880
         TabIndex        =   94
         Top             =   3600
         Width           =   720
         WordWrap        =   -1  'True
      End
      Begin VB.Label payment_detail_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Date"
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
         Left            =   9600
         TabIndex        =   93
         Top             =   3600
         Width           =   1440
      End
      Begin VB.Label payment_detail_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Number"
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
         Left            =   5880
         TabIndex        =   92
         Top             =   3600
         Width           =   1680
      End
      Begin VB.Label payment_detail_fr_label 
         AutoSize        =   -1  'True
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
         Index           =   4
         Left            =   2280
         TabIndex        =   91
         Top             =   3600
         Width           =   1320
      End
      Begin VB.Label payment_detail_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room"
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
         TabIndex        =   90
         Top             =   3600
         Width           =   480
      End
      Begin VB.Label payment_detail_fr_label 
         AutoSize        =   -1  'True
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
         TabIndex        =   89
         Top             =   3600
         Width           =   480
      End
      Begin VB.Label payment_detail_fr_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Details"
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
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label payment_detail_fr_label 
         AutoSize        =   -1  'True
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
      Begin VB.Image Image25 
         Height          =   1050
         Left            =   5520
         Picture         =   "Payment_Form.frx":3FAC1
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1050
      End
      Begin VB.Image Image17 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   720
         Picture         =   "Payment_Form.frx":497CB
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   3135
      End
      Begin VB.Image Image26 
         Height          =   735
         Left            =   720
         Picture         =   "Payment_Form.frx":49CD5
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   21615
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   720
         Picture         =   "Payment_Form.frx":4A1E0
         Stretch         =   -1  'True
         Top             =   11040
         Width           =   2775
      End
   End
End
Attribute VB_Name = "Payment_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim integer_value(5) As Integer

Dim is_pending As Boolean
Dim redundency As Boolean
Dim MsgBox_Response As Integer

Dim service_1 As String * 57
Dim service_2 As String * 57

Dim room_1 As room_class
Dim room_2 As room_class
Dim room_3 As room_class

Dim service_new As service
Dim service_temp As service
Dim service_existing As service

Dim count_payment_1 As Integer
Dim count_payment_2 As Integer
Dim count_payment_3 As Integer

Dim new_payment As payment_class
Dim new_pay As payment_class
Dim existing_pay As payment_class

Dim dis_pay As display_payment_class

Dim new_electricity As electricity
Dim temp_electricity As electricity
Dim existing_electricity As electricity


Private Sub electricity_fee_extract_Click()
    count_payment_1 = electricity_fee_count_function
    If count_payment_1 = 0 Then
        MsgBox_Response = MsgBox("     Sorry no electricity fee rate has been added till now.", vbInformation + vbOKOnly, "Rental Record")
    Else
        Open "ElectricityFee.txt" For Random As #1 Len = 78
            Get #1, count_payment_1, existing_electricity
        Close #1

        j = 0
        For i = 0 To 20 Step 4 'assigning values
            unit(i + 0).Text = existing_electricity.range_min(j)
            unit(i + 1).Text = existing_electricity.range_max(j)
            unit(i + 2).Text = existing_electricity.monthly_min(j)
            unit(i + 3).Text = existing_electricity.per_unit(j)
            j = j + 1
        Next i
    End If
End Sub

Private Sub electricity_fee_goback_Click()
    'form and frame visibility
    Main_Form.WindowState = Payment_Form.WindowState
    Main_Form.Visible = True
    
    Payment_Form.Visible = False
    Payment_Form.electricity_fee_frame.Visible = False
    
    'reset textbox values
    For i = 0 To 23 Step 4
        unit(i).Text = "min"
        unit(i + 1).Text = "max"
        unit(i + 2).Text = "monthly min"
        unit(i + 3).Text = "per unit"
    Next i
End Sub

Private Sub electricity_fee_save_Click()
    'check for default values
    For i = 0 To 20 Step 4
        If unit(i + 0).Text = "min" Then
            msgbox_reponse = MsgBox("     Please enter the valid input.", vbInformation + vbOKOnly, "Rental Record")
            unit(i + 0).SetFocus
            Exit Sub
        End If
        If unit(i + 1).Text = "max" Then
            msgbox_reponse = MsgBox("     Please enter the valid input.", vbInformation + vbOKOnly, "Rental Record")
            unit(i + 1).SetFocus
            Exit Sub
        End If
        If unit(i + 2).Text = "monthly min" Then
            msgbox_reponse = MsgBox("     Please enter the valid input.", vbInformation + vbOKOnly, "Rental Record")
            unit(i + 2).SetFocus
            Exit Sub
        End If
        If unit(i + 3).Text = "per unit" Then
            msgbox_reponse = MsgBox("     Please enter the valid input.", vbInformation + vbOKOnly, "Rental Record")
            unit(i + 3).SetFocus
            Exit Sub
        End If
    Next i
        
    'default value check in combo box
    If unit_combo(0).Text = "Year" Or unit_combo(1).Text = "Month" Then
        If unit_combo(0).Text = "Year" Then
            MsgBox_Response = MsgBox("     Please select the year first", vbInformation + vbOKOnly, "Rental Record")
            unit_combo(0).SetFocus
        Else
            MsgBox_Response = MsgBox("     Please select the year first", vbInformation + vbOKOnly, "Rental Record")
            unit_combo(1).SetFocus
        End If
        Exit Sub
    End If
    
    j = 0
    For i = 0 To 20 Step 4 'assigning values
        new_electricity.range_min(j) = Val(unit(i + 0).Text)
        new_electricity.range_max(j) = Val(unit(i + 1).Text)
        new_electricity.monthly_min(j) = Val(unit(i + 2).Text)
        new_electricity.per_unit(j) = Val(unit(i + 3).Text)
        j = j + 1
    Next i
      
    new_electricity.electricity_year(1) = 0
    new_electricity.electricity_month(1) = 0
    
    new_electricity.electricity_year(0) = Val(unit_combo(0))
    
    new_electricity.electricity_month(0) = set_month_in_integer(unit_combo(1))
    count_payment_1 = electricity_fee_count_function
    'new_electricity.electricity_month(0) = Val(unit_combo(1))
    
    Open "ElectricityFee.txt" For Random As #1 Len = 78
        If count_payment_1 = 0 Then
            Put #1, 1, new_electricity
            MsgBox_Response = MsgBox("     New electricity fee rate has been added.", vbInformation + vbOKOnly, "Rental Record")
        Else '>> presence of electricity fee
            Get #1, count_payment_1, existing_electricity
            
            If existing_electricity.electricity_year(0) > Val(unit_combo(0).Text) Then ' adding for previous date
                MsgBox_Response = MsgBox("     Sorry! you cannot add detail for this year.", vbInformation + vbOKOnly, "Rental Record")
                Close #1
                Exit Sub
                
            ElseIf existing_electricity.electricity_year(0) = Val(unit_combo(0).Text) Then 'same year
                
                If existing_electricity.electricity_month(0) > new_electricity.electricity_month(0) Then 'smaller month >> invalid month
                    MsgBox_Response = MsgBox("     Sorry! you cannot add detail for this month.", vbInformation + vbOKOnly, "Rental Record")
                    Close #1
                    Exit Sub
                
                ElseIf existing_electricity.electricity_month(0) = new_electricity.electricity_month(0) Then 'same month
                    'update value
                    Put #1, count_payment_1, new_electricity
                    MsgBox_Response = MsgBox("     Electricity fee rate has been updated.", vbInformation + vbOKOnly, "Rental Record")
                
                Else 'greater month
                    Get #1, count_payment_1, existing_electricity
                        existing_electricity.electricity_year(1) = new_electricity.electricity_year(0)
                        existing_electricity.electricity_month(1) = new_electricity.electricity_month(0)
                    Put #1, count_payment_1, existing_electricity
                    Put #1, count_payment_1 + 1, new_electricity
                    MsgBox_Response = MsgBox("     New electricity fee rate has been added.", vbInformation + vbOKOnly, "Rental Record")
                End If
            
            Else 'future year
                Get #1, count_payment_1, existing_electricity
                    existing_electricity.electricity_year(1) = new_electricity.electricity_year(0)
                    existing_electricity.electricity_month(1) = new_electricity.electricity_month(0)
                Put #1, count_payment_1, existing_electricity
                Put #1, count_payment_1 + 1, new_electricity
                MsgBox_Response = MsgBox("     New electricity fee rate has been added.", vbInformation + vbOKOnly, "Rental Record")
            End If
        End If
    Close #1
    
    'reset textbox values
    For i = 0 To 20 Step 4
        unit(i + 0) = "min"
        unit(i + 1) = "max"
        unit(i + 2) = "monthly min"
        unit(i + 3) = "per unit"
    Next i
End Sub

Private Sub fee_combo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 0 Then
            If fee_combo(0).Text <> "Year" Then fee_combo(1).SetFocus
        ElseIf Index = 1 Then
            If fee_combo(1).Text <> "Month" Then service_add_fr_save_Click
        End If
    Else
        KeyAscii = 0
    End If
End Sub

'add new service
Private Sub fee_textbox_GotFocus(Index As Integer)
    For i = 0 To 3
        If Index = i And fee_textbox(i).Text = "Enter fee" Then
            fee_textbox(i).Text = ""
            Exit Sub
        End If
    Next i
End Sub

Private Sub fee_textbox_LostFocus(Index As Integer)
    For i = 0 To 3
        If Index = i And fee_textbox(i).Text = "" Then
            fee_textbox(i).Text = "Enter fee"
            Exit Sub
        End If
    Next i
End Sub

Private Sub fee_textbox_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 46 Then
        KeyAscii = KeyAscii
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        
        For i = 0 To 2
            If Index = i And fee_textbox(i).Text <> "Enter fee" And fee_textbox(i).Text <> "" Then
                fee_textbox(i + 1).SetFocus
                Exit Sub
            End If
        Next i

        If Index = 3 Then
            If fee_textbox(3).Text <> "Enter fee" And fee_textbox(3).Text <> "" Then fee_combo(0).SetFocus
        End If
        
    ElseIf KeyAscii > 47 And KeyAscii < 58 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub issue_detail_goback_Click()
    'form and frame visibility
    Main_Form.WindowState = Payment_Form.WindowState
    Main_Form.Visible = True
    
    Payment_Form.Visible = False
    '---------dasdada
    Payment_Form.service_det_frame.Visible = False
    Payment_Form.electricity_fee_frame.Visible = False
    Payment_Form.service_add_fr.Visible = False
    Payment_Form.payment_frame.Visible = False
    Payment_Form.payment_detail_frame.Visible = False
End Sub

Private Sub payment_detail_fr_goback_Click()
    'form and frame visibility
    Main_Form.WindowState = Payment_Form.WindowState
    Main_Form.Visible = True
    Payment_Form.Visible = False
    Payment_Form.payment_detail_frame.Visible = False
End Sub

Private Sub payment_filter_combo_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub payment_filter_command_Click()
    count_payment_2 = payment_detail_count_function
    count_payment_3 = 1
    payment_detail_TextBox.Text = ""
        
    If payment_filter_check(0).Value = Checked Or payment_filter_check(1).Value = Checked Or payment_filter_check(2).Value = Checked Or payment_filter_check(3).Value = Checked Then
        payment_filter_command.Caption = "Filter : On"
        Open "PaymentDetail.txt" For Random As #2 Len = 79
            count_payment_3 = 1
            For count_payment_1 = 1 To count_payment_2 Step 1
                Get #2, count_payment_1, existing_pay
                dis_pay.serial = count_payment_3
                dis_pay.room_num = existing_pay.room_num
                dis_pay.fname = existing_pay.fname
                
                If existing_pay.mname = "Null      " Then
                    dis_pay.mname = existing_pay.lname
                    dis_pay.lname = ""
                Else
                    dis_pay.mname = existing_pay.mname
                    dis_pay.lname = existing_pay.lname
                End If
                
                dis_pay.contact = existing_pay.contact
                dis_pay.year = existing_pay.year
                dis_pay.rent_amount = existing_pay.rent_amount
                dis_pay.water_fee = existing_pay.water_fee
                dis_pay.waste_fee = existing_pay.waste_fee
                dis_pay.security_fee = existing_pay.security_fee
                dis_pay.internet_fee = existing_pay.internet_fee
                
                If existing_pay.total = 0# Then
                    dis_pay.payment_year = "-"
                    dis_pay.payment_month = "-"
                    dis_pay.payment_date = "-"
                    dis_pay.elec_unit = "-"
                    dis_pay.electricity_fee = "-"
                    dis_pay.total = "-"
                Else
                    dis_pay.payment_year = existing_pay.payment_year
                    dis_pay.payment_month = existing_pay.payment_month
                    dis_pay.payment_date = existing_pay.payment_date
                    dis_pay.elec_unit = existing_pay.elec_unit
                    dis_pay.electricity_fee = existing_pay.electricity_fee
                    dis_pay.total = existing_pay.total
                End If
                
                Call return_month_in_string(dis_pay.month, existing_pay.month)
                Call return_month_in_string(dis_pay.payment_month, existing_pay.payment_month)
                
                'applying filters
                '>>filter off
                If payment_filter_check(0).Value = Unchecked And payment_filter_check(1).Value = Unchecked And payment_filter_check(2).Value = Unchecked And payment_filter_check(3).Value = Unchecked Then
                    Call display_payment_detail_function(dis_pay)
                    
                'based on room number
                ElseIf payment_filter_check(0).Value = Checked And payment_filter_check(1).Value = Unchecked And payment_filter_check(2).Value = Unchecked And payment_filter_check(3).Value = Unchecked Then
                    If payment_filter_combo(0).Text <> "Room No" Then
                        If existing_pay.room_num = Val(payment_filter_combo(0).Text) Then
                            Call display_payment_detail_function(dis_pay)
                            count_payment_3 = count_payment_3 + 1
                        End If
                    Else
                        MsgBox_Response = MsgBox("     Please select the room number first.", vbInformation + vbOKOnly, "Rental Record")
                        count_payment_1 = count_payment_2 + 2
                    End If
                
                'based on year
                ElseIf payment_filter_check(0).Value = Unchecked And payment_filter_check(1).Value = Checked And payment_filter_check(2).Value = Unchecked And payment_filter_check(3).Value = Unchecked Then
                    If payment_filter_combo(1).Text <> "Year" Then
                        If existing_pay.year = Val(payment_filter_combo(1).Text) Then
                            Call display_payment_detail_function(dis_pay)
                            count_payment_3 = count_payment_3 + 1
                        End If
                    Else
                        MsgBox_Response = MsgBox("     Please select the year first.", vbInformation + vbOKOnly, "Rental Record")
                        count_payment_1 = count_payment_2 + 2
                    End If
                    
                'based on month
                ElseIf payment_filter_check(0).Value = Unchecked And payment_filter_check(1).Value = Unchecked And payment_filter_check(2).Value = Checked And payment_filter_check(3).Value = Unchecked Then
                    If payment_filter_combo(2).Text <> "Month" Then
                        If existing_pay.month = set_month_in_integer(payment_filter_combo(2).Text) Then
                            Call display_payment_detail_function(dis_pay)
                            count_payment_3 = count_payment_3 + 1
                        End If
                    Else
                        MsgBox_Response = MsgBox("     Please select the month.", vbInformation + vbOKOnly, "Rental Record")
                        count_payment_1 = count_payment_2 + 2
                    End If
                    
                'based on payment status
                ElseIf payment_filter_check(0).Value = Unchecked And payment_filter_check(1).Value = Unchecked And payment_filter_check(2).Value = Unchecked And payment_filter_check(3).Value = Checked Then
                    If payment_filter_option(0).Value = True Then 'paid
                        If existing_pay.total <> 0# Then
                            Call display_payment_detail_function(dis_pay)
                            count_payment_3 = count_payment_3 + 1
                        End If
                    Else  'unpaid
                        If existing_pay.total = 0# Then
                            Call display_payment_detail_function(dis_pay)
                            count_payment_3 = count_payment_3 + 1
                        End If
                    End If
                    
                'based on room number and year
                ElseIf payment_filter_check(0).Value = Checked And payment_filter_check(1).Value = Checked And payment_filter_check(2).Value = Unchecked And payment_filter_check(3).Value = Unchecked Then
                    If payment_filter_combo(0).Text <> "Room No" And payment_filter_combo(1).Text <> "Year" Then
                        If existing_pay.room_num = Val(payment_filter_combo(0).Text) And existing_pay.year = Val(payment_filter_combo(1).Text) Then
                            Call display_payment_detail_function(dis_pay)
                            count_payment_3 = count_payment_3 + 1
                        End If
                    Else
                        MsgBox_Response = MsgBox("     Please select the room number and year.", vbInformation + vbOKOnly, "Rental Record")
                        count_payment_1 = count_payment_2 + 2
                    End If
                    
                'based on room number and month
                ElseIf payment_filter_check(0).Value = Checked And payment_filter_check(1).Value = Unchecked And payment_filter_check(2).Value = Checked And payment_filter_check(3).Value = Unchecked Then
                    If payment_filter_combo(0).Text <> "Room No" And payment_filter_combo(2).Text <> "Month" Then
                        If existing_pay.room_num = Val(payment_filter_combo(0).Text) And existing_pay.month = set_month_in_integer(payment_filter_combo(2).Text) Then
                            Call display_payment_detail_function(dis_pay)
                            count_payment_3 = count_payment_3 + 1
                        End If
                    Else
                        MsgBox_Response = MsgBox("     Please select the room number, year and month.", vbInformation + vbOKOnly, "Rental Record")
                        count_payment_1 = count_payment_2 + 2
                    End If
                
                'based on room number and payament status
                ElseIf payment_filter_check(0).Value = Checked And payment_filter_check(1).Value = Unchecked And payment_filter_check(2).Value = Unchecked And payment_filter_check(3).Value = Checked Then
                    If payment_filter_combo(0).Text <> "Room No" Then
                        If existing_pay.room_num = Val(payment_filter_combo(0).Text) Then
                        
                            If payment_filter_option(0).Value = True Then
                                If existing_pay.total <> 0# Then 'paid details option
                                    Call display_payment_detail_function(dis_pay)
                                    count_payment_3 = count_payment_3 + 1
                                End If
                            ElseIf payment_filter_option(0).Value = False Then
                                If existing_pay.total = 0# Then 'unpaid details
                                    Call display_payment_detail_function(dis_pay)
                                    count_payment_3 = count_payment_3 + 1
                                End If
                            End If
                        End If
                    Else
                        MsgBox_Response = MsgBox("     Please select the room number.", vbInformation + vbOKOnly, "Rental Record")
                        count_payment_1 = count_payment_2 + 2
                    End If
                    
                'based on year and month
                ElseIf payment_filter_check(0).Value = Unchecked And payment_filter_check(1).Value = Checked And payment_filter_check(2).Value = Checked And payment_filter_check(3).Value = Unchecked Then
                    If payment_filter_combo(1).Text <> "Year" And payment_filter_combo(2).Text <> "Month" Then
                        If existing_pay.year = Val(payment_filter_combo(1).Text) And existing_pay.month = set_month_in_integer(payment_filter_combo(2).Text) Then
                            Call display_payment_detail_function(dis_pay)
                            count_payment_3 = count_payment_3 + 1
                        End If
                    Else
                        MsgBox_Response = MsgBox("     Please select the year and the month.", vbInformation + vbOKOnly, "Rental Record")
                        count_payment_1 = count_payment_2 + 2
                    End If
                     
                'based on year and rent status
                ElseIf payment_filter_check(0).Value = Unchecked And payment_filter_check(1).Value = Checked And payment_filter_check(2).Value = Unchecked And payment_filter_check(3).Value = Checked Then
                    If payment_filter_combo(1).Text <> "Year" Then
                        If existing_pay.year = Val(payment_filter_combo(1).Text) Then
                            If payment_filter_option(0).Value = True Then 'paid
                                If existing_pay.total <> 0# Then
                                    Call display_payment_detail_function(dis_pay)
                                    count_payment_3 = count_payment_3 + 1
                                End If
                            Else
                                If existing_pay.total = 0# Then 'unpaid
                                    Call display_payment_detail_function(dis_pay)
                                    count_payment_3 = count_payment_3 + 1
                                End If
                            End If
                        End If
                    Else
                        MsgBox_Response = MsgBox("     Please select the year and the month.", vbInformation + vbOKOnly, "Rental Record")
                        count_payment_1 = count_payment_2 + 1
                    End If
                           
                'based on month and rent status
                ElseIf payment_filter_check(0).Value = Unchecked And payment_filter_check(1).Value = Unchecked And payment_filter_check(2).Value = Checked And payment_filter_check(3).Value = Checked Then
                    If payment_filter_combo(2).Text <> "Month" Then
                        If existing_pay.month = set_month_in_integer(payment_filter_combo(2).Text) Then
                            If payment_filter_option(0).Value = True Then 'paid
                                If existing_pay.total <> 0# Then
                                    Call display_payment_detail_function(dis_pay)
                                    count_payment_3 = count_payment_3 + 1
                                End If
                            Else  'unpaid
                                If existing_pay.total = 0# Then
                                    Call display_payment_detail_function(dis_pay)
                                    count_payment_3 = count_payment_3 + 1
                                End If
                            End If
                        End If
                    Else
                        MsgBox_Response = MsgBox("     Please select the month.", vbInformation + vbOKOnly, "Rental Record")
                        count_payment_1 = count_payment_2 + 2
                    End If
                
                'based on room number, year and month
                ElseIf payment_filter_check(0).Value = Checked And payment_filter_check(1).Value = Checked And payment_filter_check(2).Value = Checked And payment_filter_check(3).Value = Unchecked Then
                    If payment_filter_combo(0).Text <> "Room No" And payment_filter_combo(1).Text <> "Year" And payment_filter_combo(2).Text <> "Month" Then
                        If existing_pay.room_num = Val(payment_filter_combo(0).Text) And existing_pay.year = Val(payment_filter_combo(1).Text) And existing_pay.month = set_month_in_integer(payment_filter_combo(2).Text) Then
                            Call display_payment_detail_function(dis_pay)
                            count_payment_3 = count_payment_3 + 1
                        End If
                    Else
                        MsgBox_Response = MsgBox("     Make sure you selected the room number, year and month.", vbInformation + vbOKOnly, "Rental Record")
                        count_payment_1 = count_payment_2 + 2
                    End If
                
                'based on room number, year and payment status
                ElseIf payment_filter_check(0).Value = Checked And payment_filter_check(1).Value = Checked And payment_filter_check(2).Value = Unchecked And payment_filter_check(3).Value = Checked Then
                    If payment_filter_combo(0).Text <> "Room No" And payment_filter_combo(1).Text <> "Year" Then
                        If existing_pay.room_num = Val(payment_filter_combo(0).Text) And existing_pay.year = Val(payment_filter_combo(1).Text) Then
                            If payment_filter_option(0).Value = True Then 'paid
                                If existing_pay.total <> 0# Then
                                    Call display_payment_detail_function(dis_pay)
                                    count_payment_3 = count_payment_3 + 1
                                End If
                            Else
                                If existing_pay.total = 0# Then 'unpaid
                                    Call display_payment_detail_function(dis_pay)
                                    count_payment_3 = count_payment_3 + 1
                                End If
                            End If
                        End If
                    Else
                        MsgBox_Response = MsgBox("     Make sure you selected the room number and year.", vbInformation + vbOKOnly, "Rental Record")
                        count_payment_1 = count_payment_2 + 2
                    End If
                
                'based on room number. month and the payment status
                ElseIf payment_filter_check(0).Value = Checked And payment_filter_check(1).Value = Unchecked And payment_filter_check(2).Value = Checked And payment_filter_check(3).Value = Checked Then
                    If payment_filter_combo(0).Text <> "Room No" And payment_filter_combo(2).Text <> "Month" Then
                        If existing_pay.room_num = Val(payment_filter_combo(0).Text) And existing_pay.month = set_month_in_integer(payment_filter_combo(2).Text) Then
                            If payment_filter_option(0).Value = True Then 'paid
                                If existing_pay.total <> 0# Then
                                    Call display_payment_detail_function(dis_pay)
                                    count_payment_3 = count_payment_3 + 1
                                End If
                            Else
                                If existing_pay.total = 0# Then 'unpaid
                                    Call display_payment_detail_function(dis_pay)
                                    count_payment_3 = count_payment_3 + 1
                                End If
                            End If
                        End If
                    Else
                        MsgBox_Response = MsgBox("     Make sure you selected the room number and year.", vbInformation + vbOKOnly, "Rental Record")
                        count_payment_1 = count_payment_2 + 2
                    End If
                
                'based on year, month and the payment status
                ElseIf payment_filter_check(0).Value = Unchecked And payment_filter_check(1).Value = Checked And payment_filter_check(2).Value = Checked And payment_filter_check(3).Value = Checked Then
                    If payment_filter_combo(1).Text <> "Year" And payment_filter_combo(2).Text <> "Month" Then
                        If existing_pay.year = Val(payment_filter_combo(1).Text) And existing_pay.month = set_month_in_integer(payment_filter_combo(2).Text) Then
                            If payment_filter_option(0).Value = True Then 'paid
                                If existing_pay.total <> 0# Then
                                    Call display_payment_detail_function(dis_pay)
                                    count_payment_3 = count_payment_3 + 1
                                End If
                            Else
                                If existing_pay.total = 0# Then 'unpaid
                                    Call display_payment_detail_function(dis_pay)
                                    count_payment_3 = count_payment_3 + 1
                                End If
                            End If
                        End If
                    Else
                        MsgBox_Response = MsgBox("     Make sure you selected the room number and year.", vbInformation + vbOKOnly, "Rental Record")
                        count_payment_1 = count_payment_2 + 2
                        Call fill_payment_details
                    End If
                
                'based on room number, year, month and payment status
                ElseIf payment_filter_check(0).Value = Checked And payment_filter_check(1).Value = Checked And payment_filter_check(2).Value = Checked And payment_filter_check(3).Value = Checked Then
                    If payment_filter_combo(0).Text <> "Room No" And payment_filter_combo(1).Text <> "Year" And payment_filter_combo(2).Text <> "Month" Then
                        If existing_pay.room_num = Val(payment_filter_combo(0).Text) And existing_pay.year = Val(payment_filter_combo(1).Text) And existing_pay.month = set_month_in_integer(payment_filter_combo(2).Text) Then
                            If payment_filter_option(0).Value = True Then 'paid
                                If existing_pay.total <> 0# Then
                                    Call display_payment_detail_function(dis_pay)
                                    count_payment_3 = count_payment_3 + 1
                                End If
                            Else
                                If existing_pay.total = 0# Then 'unpaid
                                    Call display_payment_detail_function(dis_pay)
                                    count_payment_3 = count_payment_3 + 1
                                End If
                            End If
                        End If
                    Else
                        MsgBox_Response = MsgBox("     Make sure you selected the room number, year and month.", vbInformation + vbOKOnly, "Rental Record")
                        count_payment_1 = count_payment_2 + 2
                    End If
                End If
            Next count_payment_1
        Close #2
    Else
        payment_filter_command.Caption = "Filter : Off"
        Call fill_payment_details
    End If
End Sub


Private Sub payment_goback_1_Click()
    'reset values
    payment_room_combo.Text = "Room number"
    payment_unit_TextBox.Text = "Enter electricity unit"
    
    'form and frame visibility
    Main_Form.WindowState = Payment_Form.WindowState
    Main_Form.Visible = True
    Payment_Form.Visible = False
    payment_frame.Visible = False
End Sub

Private Sub payment_goback_2_Click()
    'payment_goback_1_Click
    
    'frame visibility
    payment_selection_fr.Visible = True
    payment_confirm_fr.Visible = False
End Sub

Private Sub payment_next_command_Click()
    If payment_next_command.Caption = "Check" Then
        'check for default values
        If payment_room_combo.Text = "Room number" Then
            MsgBox_Response = MsgBox("     Please select the room number first.", vbInformation + vbOKOnly, "Rental Record")
            payment_room_combo.SetFocus
            Exit Sub
        End If
        
        'get tenant detail from room file
        count_payment_2 = room_detail_count_function
        Open "RoomDetail.txt" For Random As #1 Len = 112
            For count_payment_1 = 1 To count_payment_2 Step 1
                Get #1, count_payment_1, room_1
                If room_1.room_number = Val(payment_room_combo.Text) Then count_payment_1 = count_payment_2 + 2
            Next count_payment_1
        Close #1
        
        'check for the presence of pending payment detail for a provided room number in payment detail file
        is_pending = False
        is_pending = pending_payment_check(room_1.room_number, room_1.tenant_fname, room_1.tenant_mname, room_1.tenant_lname)
        
        'count number of details present in a file
        
        i = payment_detail_count_function
        
        'status : true if left to be paid
        
        If is_pending = True And i > 0 Then
            payment_next_command.Caption = "Proceed"
            payment_unit_TextBox.Visible = True
            payment_fr_label(3).Visible = True
            payment_fr_label(30).Visible = True
        Else
            MsgBox_Response = MsgBox("     No pending rent for this room number.", vbInformation + vbOKOnly, "Rental Record")
        End If
    ElseIf payment_next_command.Caption = "Proceed" Then
        'check for electricity unit
        If payment_unit_TextBox.Text <> "" And payment_unit_TextBox.Text <> "Enter electricity unit" Then
            'extract date form pending payment file
            Call get_pending_detail(new_pay, room_1.room_number, room_1.tenant_fname, room_1.tenant_mname, room_1.tenant_lname)
            
            payment_fr_label(10).Caption = new_pay.room_num
            payment_fr_label(11).Caption = new_pay.fname
            payment_fr_label(33).Caption = new_pay.mname
            payment_fr_label(34).Caption = new_pay.lname
            
            If new_pay.mname <> "Null      " Then
                payment_fr_label(43).Caption = new_pay.fname & new_pay.mname & new_pay.lname
            Else
                payment_fr_label(43).Caption = new_pay.fname & new_pay.lname
            End If
            
            payment_fr_label(12).Caption = new_pay.contact
            payment_fr_label(22).Caption = new_pay.year
            
            
            payment_fr_label(42).Caption = new_pay.month
            
            Dim temp_string As String
            Call return_month_in_string(temp_string, new_pay.month)
            
            payment_fr_label(22).Caption = CStr(new_pay.year) + " " + temp_string
            
            payment_fr_label(23).Caption = new_pay.rent_amount
            payment_fr_label(24).Caption = new_pay.water_fee
            payment_fr_label(25).Caption = new_pay.waste_fee
            payment_fr_label(26).Caption = new_pay.security_fee
            payment_fr_label(27).Caption = new_pay.internet_fee
            payment_fr_label(28).Caption = Val(payment_unit_TextBox.Text)
            payment_fr_label(29).Caption = new_pay.electricity_fee
            payment_fr_label(31).Caption = new_pay.total
            
            'checkbox values
            For i = 0 To 4
                payment_service_check(i).Enabled = True
            Next i
            
            For i = 1 To 4
                payment_service_check(i).Value = Checked
            Next i
            
            service_2 = "Water, Waste Management, Electricity, Security, Internet"
            
            If room_1.service_provided = service_2 Then
                payment_service_check(0).Value = Checked
            Else
                payment_service_check(0).Value = Unchecked
            End If
            
            For i = 0 To 4
                payment_service_check(i).Enabled = False
            Next i
            
            Dim ele_unit As Integer
            ele_unit = Val(payment_unit_TextBox.Text)

            new_pay.electricity_fee = calculate_electricity_fee(ele_unit, new_pay.year, new_pay.month)
            
            payment_fr_label(29).Caption = new_pay.electricity_fee
            
            'totaling
            payment_fr_label(31).Caption = Val(payment_fr_label(23).Caption) + Val(payment_fr_label(24).Caption) + Val(payment_fr_label(25).Caption) + Val(payment_fr_label(26).Caption) + Val(payment_fr_label(27).Caption) + Val(payment_fr_label(29).Caption) + Val(payment_fr_label(30).Caption)
            
            payment_confirm_fr.Visible = True
            payment_selection_fr.Visible = False
        Else
            MsgBox_Response = MsgBox("     Please enter the electricity unit first.", vbInformation + vbOKOnly, "Rental Record")
            payment_unit_TextBox.SetFocus
        End If
    End If
End Sub

Private Sub payment_pay_Click()
    'gather details
    existing_pay.room_num = Val(payment_fr_label(10).Caption)
    existing_pay.fname = payment_fr_label(11).Caption
    existing_pay.mname = payment_fr_label(33).Caption
    existing_pay.lname = payment_fr_label(34).Caption
    existing_pay.year = Val(payment_fr_label(22).Caption)
    existing_pay.month = Val(payment_fr_label(42).Caption)
    existing_pay.elec_unit = Val(payment_fr_label(28).Caption)
    existing_pay.electricity_fee = Val(payment_fr_label(29).Caption)
    existing_pay.payment_year = return_year
    existing_pay.payment_month = return_month
    existing_pay.payment_date = return_day
    existing_pay.total = Val(payment_fr_label(31).Caption)
    
    count_payment_2 = payment_detail_count_function
    Open "PaymentDetail.txt" For Random As #1 Len = 79
        For count_payment_1 = 1 To count_payment_2 Step 1
            Get #1, count_payment_1, new_payment
            If new_payment.room_num = existing_pay.room_num Then
                If new_payment.fname = existing_pay.fname And new_payment.mname = existing_pay.mname And new_payment.lname = existing_pay.lname Then
                    If new_payment.year = existing_pay.year And new_payment.month = existing_pay.month Then
                        new_payment.payment_year = existing_pay.payment_year
                        new_payment.payment_month = existing_pay.payment_month
                        new_payment.payment_date = existing_pay.payment_date
                        new_payment.elec_unit = existing_pay.elec_unit
                        new_payment.electricity_fee = existing_pay.electricity_fee
                        new_payment.is_paid = True
                        
                        new_payment.total = existing_pay.total
                        
                        Put #1, count_payment_1, new_payment
                        MsgBox_Response = MsgBox("          Payment done!", vbInformation + vbOKOnly, "Rental Record")
                        payment_selection_fr.Visible = True
                        payment_confirm_fr.Visible = False
                    End If
                End If
            End If
        Next count_payment_1
    Close #1
End Sub

Private Sub payment_room_combo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If payment_room_combo.Text <> "Room number" Then
            If payment_next_command.Caption = "Check" Then
                payment_next_command_Click
            Else
                payment_unit_TextBox.SetFocus
            End If
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub payment_unit_TextBox_GotFocus()
    If payment_unit_TextBox.Text = "Enter electricity unit" Then payment_unit_TextBox.Text = ""
End Sub

Private Sub payment_unit_TextBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If payment_unit_TextBox.Text <> "" And payment_unit_TextBox.Text <> "Enter electricity unit" Then payment_next_command_Click
    ElseIf KeyAscii = 46 Or KeyAscii = 8 Then
        KeyAscii = KeyAscii
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    Else
        KeyAscii = KeyAscii
    End If
End Sub

Private Sub payment_unit_TextBox_LostFocus()
    If payment_unit_TextBox.Text = "" Then payment_unit_TextBox.Text = "Enter electricity unit"
End Sub

Private Sub service_add_fr_goback_Click()
    'reset textbox values
    For i = 0 To 3
        fee_textbox(i).Text = "Enter fee"
        service_add_fr_label(9 + i).Caption = "Rs.xxx"
    Next i
    
    'form and frame visibility
    Main_Form.WindowState = Payment_Form.WindowState
    Main_Form.Visible = True
    
    Payment_Form.Visible = False
    Payment_Form.service_add_fr.Visible = False
End Sub

'save command
Private Sub service_add_fr_save_Click()
    'check for default values
    If fee_textbox(0).Text = "Enter fee" Or fee_textbox(1).Text = "Enter fee" Or fee_textbox(2).Text = "Enter fee" Or fee_textbox(3).Text = "Enter fee" Then
        If fee_textbox(0).Text = "Enter fee" Then
            MsgBox_Response = MsgBox("     Please enter the drinking water fee first.", vbInformation + vbOKOnly, "Rental Record")
            fee_textbox(0).SetFocus
            Exit Sub
        ElseIf fee_textbox(1).Text = "Enter fee" Then
            MsgBox_Response = MsgBox("     Please enter the waste management fee first.", vbInformation + vbOKOnly, "Rental Record")
            fee_textbox(1).SetFocus
            Exit Sub
        ElseIf fee_textbox(2).Text = "Enter fee" Then
            MsgBox_Response = MsgBox("     Please enter the security fee first.", vbInformation + vbOKOnly, "Rental Record")
            fee_textbox(2).SetFocus
            Exit Sub
        ElseIf fee_textbox(3).Text = "Enter fee" Then
            MsgBox_Response = MsgBox("     Please enter the internet fee first.", vbInformation + vbOKOnly, "Rental Record")
            fee_textbox(3).SetFocus
            Exit Sub
        End If
        Exit Sub
    End If
    
    'check for default values in combo box
    If fee_combo(0).Text = "Year" Or fee_combo(1).Text = "Month" Then
        If fee_combo(0).Text = "Year" Then
            MsgBox_Response = MsgBox("     Please select the year first", vbInformation + vbOKOnly, "Rental Record")
            fee_combo(0).SetFocus
        Else
            MsgBox_Response = MsgBox("     Please select the year month", vbInformation + vbOKOnly, "Rental Record")
            fee_combo(1).SetFocus
        End If
        Exit Sub
    End If
    
    'copy values
    For i = 0 To 3
        service_new.fee(i) = Val(fee_textbox(i).Text)
    Next i
    
    service_new.service_year(0) = Val(fee_combo(0).Text)
    service_new.service_year(1) = 0
    service_new.service_month(1) = 0
    
    'setting a date >> month
    service_new.service_month(0) = set_month_in_integer(fee_combo(1).Text)
    
    count_payment_1 = service_count_function
        
    Open "ServiceFee.txt" For Random As #1 Len = 24
        If count_payment_1 = 0 Then 'adding service fee for the first time
            Put #1, 1, service_new
            
            For i = 0 To 3 Step 1
                service_add_fr_label(9 + i).Caption = "Rs." + CStr(service_new.fee(i))
            Next i
        
            MsgBox_Response = MsgBox("     New service fee added successfully", vbInformation + vbOKOnly, "Rental Record")
        Else
            Get #1, count_payment_1, service_temp
            
            'check if the new date is smaller than the previsouly set date
            If service_temp.service_year(0) > service_new.service_year(0) Then
                MsgBox_Response = MsgBox("     Please select the valid year.", vbInformation + vbOKOnly, "Rental Record")
                Close #1
                Exit Sub
            ElseIf service_temp.service_year(0) = service_new.service_year(0) Then
                If service_temp.service_month(0) = service_new.service_month(0) Then
                    'update instead
                    Get #1, count_payment_1, service_existing
                        For i = 0 To 3
                            service_existing.fee(i) = Val(fee_textbox(i).Text)
                        Next i
                    Put #1, count_payment_1, service_existing
                    
                    MsgBox_Response = MsgBox("     Service fee updated instead.", vbInformation + vbOKOnly, "Rental Record")
                ElseIf service_temp.service_month(0) > service_new.service_month(0) Then
                    MsgBox_Response = MsgBox("     Please select the valid month.", vbInformation + vbOKOnly, "Rental Record")
                    Close #1
                    Exit Sub
                Else
                    Get #1, count_payment_1, service_existing
                        service_existing.service_year(1) = service_new.service_year(0)
                        service_existing.service_month(1) = service_new.service_month(0)
                    Put #1, count_payment_1, service_existing
                    Put #1, count_payment_1 + 1, service_new
                    MsgBox_Response = MsgBox("     New service fee added.", vbInformation + vbOKOnly, "Rental Record")
                End If
            Else ' new date >> date & month
                Get #1, count_payment_1, service_existing
                    service_existing.service_year(1) = service_new.service_year(0)
                    service_existing.service_month(1) = service_new.service_month(0)
                Put #1, count_payment_1, service_existing
                Put #1, count_payment_1 + 1, service_new
                
                MsgBox_Response = MsgBox("     New service fee added.", vbInformation + vbOKOnly, "Rental Record")
            End If
            
            'update label caption
            For i = 0 To 3
                service_add_fr_label(9 + i).Caption = "Rs." + CStr(service_new.fee(i))
            Next i
        End If
On Error GoTo ErrorHandler
    Close #1
ErrorHandler:
    
    'reset textbox values
    For i = 0 To 3
        fee_textbox(i).Text = "Enter fee"
    Next i
    
    fee_combo(0).Text = "Year"
    fee_combo(1).Text = "Month"
End Sub

'unload form
Private Sub Form_Unload(Cancel As Integer)
    Unload Main_Form
    Unload Issue_Form
    Unload Tenant_Form
End Sub

Private Sub service_detail_fr_goback_Click()
    'form and frame visibility
    Main_Form.WindowState = Payment_Form.WindowState
    Main_Form.Visible = True
    Payment_Form.Visible = False
    service_add_fr.Visible = False
    service_det_frame.Visible = False
End Sub

Private Sub service_filter_combo_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub unit_combo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 0 Then
            unit_combo(1).SetFocus
        ElseIf Index = 1 Then
            electricity_fee_save_Click
        End If
    Else
        KeyAscii = 0
    End If
End Sub


Private Sub unit_GotFocus(Index As Integer)
    For i = 0 To 20 Step 4
        If Index = i + 0 And unit(i + 0).Text = "min" Then
            unit(i + 0).Text = ""
            Exit Sub
        End If
        If Index = i + 1 And unit(i + 1).Text = "max" Then
            unit(i + 1).Text = ""
            Exit Sub
        End If
        If Index = i + 2 And unit(i + 2).Text = "monthly min" Then
            unit(i + 2).Text = ""
            Exit Sub
        End If
        If i < 21 And Index = i + 3 And unit(i + 3).Text = "per unit" Then
            unit(i + 3).Text = ""
            Exit Sub
        End If
    Next i
End Sub

Private Sub unit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 8 Then
        KeyAscii = KeyAscii
    ElseIf KeyAscii = 13 Then 'enter key press
        KeyAscii = 0
        
        For i = 0 To 22
            If Index = i And unit(i).Text <> "" Then
                unit(i + 1).SetFocus
                i = 23
            End If
        Next i
        
        If Index = 23 And unit(23).Text <> "" Then unit_combo(0).SetFocus
    ElseIf KeyAscii = 46 Then 'decimal point
        If Index = 0 Or Index = 4 Or Index = 8 Or Index = 12 Or Index = 16 Or Index = 20 Or Index = 1 Or Index = 5 Or Index = 9 Or Index = 13 Or Index = 17 Or Index = 21 Then
            KeyAscii = 0
        Else
            KeyAscii = KeyAscii
        End If
    ElseIf KeyAscii > 47 And KeyAscii < 58 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub unit_LostFocus(Index As Integer)
    For i = 0 To 20 Step 4
        If Index = i + 0 And unit(i + 0).Text = "" Then
            unit(i + 0).Text = "min"
            Exit Sub
        End If
        If Index = i + 1 And unit(i + 1).Text = "" Then
            unit(i + 1).Text = "max"
            Exit Sub
        End If
        If Index = i + 2 And unit(i + 2).Text = "" Then
            unit(i + 2).Text = "monthly min"
            Exit Sub
        End If
        If i < 21 And Index = i + 3 And unit(i + 3).Text = "" Then
            unit(i + 3).Text = "per unit"
            Exit Sub
        End If
    Next i
End Sub
