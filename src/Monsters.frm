VERSION 5.00
Begin VB.Form Monsters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monsters"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   10575
   Begin VB.Frame Frame13 
      Caption         =   "Spell Power"
      Height          =   975
      Left            =   7320
      TabIndex        =   85
      Top             =   3840
      Width           =   3015
      Begin VB.TextBox txtSpellpower 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         MaxLength       =   3
         TabIndex        =   86
         Top             =   390
         Width           =   2535
      End
   End
   Begin VB.Frame Frame12 
      Height          =   1335
      Left            =   7320
      TabIndex        =   68
      Top             =   4920
      Width           =   3135
   End
   Begin VB.Frame Frame11 
      Caption         =   "Counter Attack Chance"
      Height          =   885
      Left            =   5760
      TabIndex        =   63
      Top             =   6360
      Width           =   4695
      Begin VB.OptionButton optCounter 
         Caption         =   "1 / 16"
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   2520
         TabIndex        =   67
         Top             =   255
         Width           =   855
      End
      Begin VB.OptionButton optCounter 
         Caption         =   "1 / 32"
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   3720
         TabIndex        =   66
         Top             =   255
         Width           =   855
      End
      Begin VB.OptionButton optCounter 
         Caption         =   "1 / 8"
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   1320
         TabIndex        =   65
         Top             =   255
         Width           =   735
      End
      Begin VB.OptionButton optCounter 
         Caption         =   "1 / 4"
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   64
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Double Attack Chance"
      Height          =   885
      Left            =   5760
      TabIndex        =   58
      Top             =   7410
      Width           =   4695
      Begin VB.OptionButton optDouble 
         Caption         =   "1 / 4"
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   62
         Top             =   255
         Width           =   855
      End
      Begin VB.OptionButton optDouble 
         Caption         =   "1 / 8"
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   1320
         TabIndex        =   61
         Top             =   255
         Width           =   735
      End
      Begin VB.OptionButton optDouble 
         Caption         =   "1 / 32"
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   3720
         TabIndex        =   60
         Top             =   255
         Width           =   855
      End
      Begin VB.OptionButton optDouble 
         Caption         =   "1 / 16"
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   2520
         TabIndex        =   59
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Critical Attack Chance"
      Height          =   1935
      Left            =   120
      TabIndex        =   55
      Top             =   6360
      Width           =   5535
      Begin VB.OptionButton optCritical 
         Caption         =   "Muddle"
         Enabled         =   0   'False
         Height          =   495
         Index           =   12
         Left            =   1560
         TabIndex        =   84
         Top             =   1365
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "Slow"
         Enabled         =   0   'False
         Height          =   495
         Index           =   13
         Left            =   2600
         TabIndex        =   83
         Top             =   1365
         Width           =   735
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "Silence"
         Enabled         =   0   'False
         Height          =   495
         Index           =   15
         Left            =   4560
         TabIndex        =   82
         Top             =   1365
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "None"
         Enabled         =   0   'False
         Height          =   495
         Index           =   8
         Left            =   1560
         TabIndex        =   81
         Top             =   990
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "Stun"
         Enabled         =   0   'False
         Height          =   495
         Index           =   11
         Left            =   4560
         TabIndex        =   80
         Top             =   990
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "1 / 4"
         Enabled         =   0   'False
         Height          =   495
         Index           =   6
         Left            =   1560
         TabIndex        =   79
         Top             =   615
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "1 / 8"
         Enabled         =   0   'False
         Height          =   495
         Index           =   4
         Left            =   2600
         TabIndex        =   78
         Top             =   615
         Width           =   735
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "1 / 32"
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   4560
         TabIndex        =   77
         Top             =   615
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "1 / 16"
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   3520
         TabIndex        =   76
         Top             =   615
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "1 / 4"
         Enabled         =   0   'False
         Height          =   495
         Index           =   7
         Left            =   1560
         TabIndex        =   75
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "1 / 8"
         Enabled         =   0   'False
         Height          =   495
         Index           =   5
         Left            =   2600
         TabIndex        =   74
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "1 / 32"
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   4560
         TabIndex        =   73
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "1 / 16"
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   3520
         TabIndex        =   72
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "Poison"
         Enabled         =   0   'False
         Height          =   495
         Index           =   9
         Left            =   2600
         TabIndex        =   71
         Top             =   990
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "MP Drain"
         Enabled         =   0   'False
         Height          =   495
         Index           =   14
         Left            =   3520
         TabIndex        =   70
         Top             =   1365
         Width           =   975
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "Sleep"
         Enabled         =   0   'False
         Height          =   495
         Index           =   10
         Left            =   3520
         TabIndex        =   69
         Top             =   990
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "125% Damage"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "150% Damage"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame8 
      Height          =   735
      Left            =   10200
      TabIndex        =   52
      Top             =   7920
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox txtSpecial 
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Special Ability"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Other Resist - Select Both For Immunity"
      Height          =   855
      Left            =   7320
      TabIndex        =   49
      Top             =   2880
      Width           =   3135
      Begin VB.CheckBox chkMajor 
         Caption         =   "Major"
         Height          =   255
         Left            =   1680
         TabIndex        =   27
         Tag             =   "2"
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkMinor 
         Caption         =   "Minor"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Tag             =   "2"
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Battle Sprite"
      Height          =   855
      Left            =   3840
      TabIndex        =   46
      Top             =   5400
      Width           =   3375
      Begin VB.TextBox txtBattlePalette 
         Height          =   285
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   29
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtBattleModel 
         Height          =   285
         Left            =   840
         MaxLength       =   3
         TabIndex        =   28
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Palette"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   48
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Resistance - Select Both For Weakness"
      Height          =   2655
      Left            =   7320
      TabIndex        =   45
      Top             =   120
      Width           =   3135
      Begin VB.CheckBox chk50Fire 
         Caption         =   "50% Fire"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Tag             =   "128"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CheckBox chk25Fire 
         Caption         =   "25% Fire"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Tag             =   "64"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CheckBox chk50Ice 
         Caption         =   "50% Ice"
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Tag             =   "32"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CheckBox chk25Ice 
         Caption         =   "25% Ice"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Tag             =   "16"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CheckBox chk50Lightning 
         Caption         =   "50% Lightning"
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Tag             =   "8"
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox chk25Lightning 
         Caption         =   "25% Lightning"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Tag             =   "4"
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox chk50Wind 
         Caption         =   "50% Wind"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Tag             =   "2"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chk25Wind 
         Caption         =   "25% Wind"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Tag             =   "1"
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Spells"
      Height          =   2415
      Left            =   3840
      TabIndex        =   37
      Top             =   2880
      Width           =   3375
      Begin VB.ComboBox cboSpell 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1680
         Width           =   2895
      End
      Begin VB.ListBox lstSpells 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Inventory"
      Height          =   2655
      Left            =   3840
      TabIndex        =   36
      Top             =   120
      Width           =   3375
      Begin VB.ListBox lstItems 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   2895
      End
      Begin VB.CheckBox chkEquipped 
         Alignment       =   1  'Right Justify
         Caption         =   "Equipped"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2160
         Width           =   2775
      End
      Begin VB.ComboBox cboItems 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1560
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5295
      Left            =   120
      TabIndex        =   34
      Top             =   960
      Width           =   3615
      Begin VB.TextBox txtGold 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   9
         Top             =   3720
         Width           =   2295
      End
      Begin VB.ComboBox cboMoveType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   4680
         Width           =   2295
      End
      Begin VB.TextBox txtMove 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   10
         Top             =   4200
         Width           =   2295
      End
      Begin VB.TextBox txtHPTimes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtLevel 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtAGI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   8
         Top             =   3240
         Width           =   2295
      End
      Begin VB.TextBox txtDEF 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   7
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox txtATK 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   6
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox txtMP 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   5
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtHP 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   3
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   30
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label13 
         Caption         =   "Gold"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "MV Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   4730
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Move"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "HP * 256"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Agility"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Defense"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Attack"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "MP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "HP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   11520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Top             =   240
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select"
      Height          =   735
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton cmdRename 
         Caption         =   "Rename"
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cboName 
         Height          =   315
         ItemData        =   "Monsters.frx":0000
         Left            =   240
         List            =   "Monsters.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.TextBox txtNameData 
      Enabled         =   0   'False
      Height          =   1365
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   31
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Monsters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Loading As Boolean
Private Address As Long
Private NameAddress As Long

Private Sub cboMoveType_Click()
  txtNameData.Visible = False
  
  RomDump(Address + 49) = cboMoveType.ItemData(cboMoveType.ListIndex)
End Sub

Private Sub chk25Ice_Click()
 
 Dim Sum As Byte
  
 txtNameData.Visible = False
 
 If Loading = True Then
  Exit Sub
 End If
 
 If chk25Fire.Value = vbChecked Then
  Sum = Sum + chk25Fire.Tag
 End If
 If chk50Fire.Value = vbChecked Then
  Sum = Sum + chk50Fire.Tag
 End If
 If chk25Ice.Value = vbChecked Then
  Sum = Sum + chk25Ice.Tag
 End If
 If chk50Ice.Value = vbChecked Then
  Sum = Sum + chk50Ice.Tag
 End If
 If chk25Wind.Value = vbChecked Then
  Sum = Sum + chk25Wind.Tag
 End If
 If chk50Wind.Value = vbChecked Then
  Sum = Sum + chk50Wind.Tag
 End If
 If chk25Lightning.Value = vbChecked Then
  Sum = Sum + chk25Lightning.Tag
 End If
 If chk50Lightning.Value = vbChecked Then
  Sum = Sum + chk50Lightning.Tag
 End If
 
 RomDump(Address + 27) = Sum
End Sub

Private Sub chk25Fire_Click()
 Call chk25Ice_Click
 txtNameData.Visible = False
End Sub

Private Sub chk25Lightning_Click()
 Call chk25Ice_Click
 txtNameData.Visible = False
End Sub

Private Sub chk25Wind_Click()
 txtNameData.Visible = False
 Call chk25Ice_Click
End Sub

Private Sub chk50Fire_Click()
 Call chk25Ice_Click
 txtNameData.Visible = False
End Sub

Private Sub chk50Ice_Click()
 Call chk25Ice_Click
 txtNameData.Visible = False
End Sub

Private Sub chk50Lightning_Click()
 Call chk25Ice_Click
 txtNameData.Visible = False
End Sub

Private Sub chk50Wind_Click()
 Call chk25Ice_Click
 txtNameData.Visible = False
End Sub

Private Sub chkEquipped_Click()
 txtNameData.Visible = False
 txtNameData.Visible = False
 
 Call cboItems_Click
End Sub

Private Sub cboItems_Click()
  
 txtNameData.Visible = False
 
 If Loading = True Or lstItems.ListIndex = -1 Then
  Exit Sub
 End If
 
 Address = cboName.ItemData(cboName.ListIndex)
 
 If chkEquipped.Value = vbChecked Then
  
  lstItems.List(lstItems.ListIndex) = cboItems.Text & " (Equipped)"
  lstItems.ItemData(lstItems.ListIndex) = cboItems.ItemData(cboItems.ListIndex) + &H80
  
  RomDump(Address + 33 + 2 * lstItems.ListIndex) = cboItems.ItemData(cboItems.ListIndex) + &H80
 
 Else
  
  lstItems.List(lstItems.ListIndex) = cboItems.Text
  lstItems.ItemData(lstItems.ListIndex) = cboItems.ItemData(cboItems.ListIndex)
  
  RomDump(Address + 33 + 2 * lstItems.ListIndex) = cboItems.ItemData(cboItems.ListIndex)
  
 End If
 
'     lstItems.AddItem ItemName(RomDump(Address + 33 + Index) - &H80) & " (Equipped)"
'     lstItems.ItemData(lstItems.ListCount - 1) = RomDump(Address + 33 + Index)
'   Else
'     lstItems.AddItem ItemName(RomDump(Address + 33 + Index))
'     lstItems.ItemData(lstItems.ListCount - 1) = RomDump(Address + 33 + Index)
 
End Sub

Private Sub cboName_Click()
  
  Dim Index As Long
  Dim SubIndex As Long
  Dim Leng As Byte
  
  txtNameData.Visible = False
  EnableTextboxes True, Me
   
  Loading = True
  
  Text1.Text = ""
  lstSpells.Clear
  lstItems.Clear
  
  chk25Fire.Value = vbUnchecked
  chk25Ice.Value = vbUnchecked
  chk25Lightning.Value = vbUnchecked
  chk25Wind.Value = vbUnchecked
  chk50Fire.Value = vbUnchecked
  chk50Ice.Value = vbUnchecked
  chk50Lightning.Value = vbUnchecked
  chk50Wind.Value = vbUnchecked
  
  chkMinor.Value = vbUnchecked
  chkMajor.Value = vbUnchecked
  
  For Index = 0 To 55
   Text1.Text = Text1.Text & Align(RomDump(cboName.ItemData(cboName.ListIndex) + Index)) & RomDump(cboName.ItemData(cboName.ListIndex) + Index) & " "
   
   If Index / 5 = Fix(Index / 5) Then
    Text1.Text = Text1.Text & vbNewLine
   End If
  Next Index
  
  Address = cboName.ItemData(cboName.ListIndex)
'''  NameAddress = GetNameAddress(cboName.ListIndex)
'''
'''  Leng = RomDump(NameAddress - 1)
'''  txtName.MaxLength = Leng
'''  txtNameData.Text = ""
'''
'''  For Index = 0 To Leng - 1
'''   If Index <> Leng - 1 Then
'''   txtNameData.Text = txtNameData.Text & RomDump(NameAddress + Index) & vbNewLine
'''   Else
'''   txtNameData.Text = txtNameData.Text & RomDump(NameAddress + Index)
'''   End If
'''  Next
'''
'''
'''  ' This bloc loads in the raw number data used for these names
'''  Dim NameArray() As String
'''  NameArray = Split(txtNameData.Text, vbNewLine)
'''  txtName.Text = ""
'''  For Index = 0 To UBound(NameArray())
'''   txtName.Text = txtName.Text & Chr(CLng(NameArray(Index)))
'''  Next Index
  

  For SubIndex = 0 To 3
  
  Index = 0
  Do While SpellCode(Index) <> RomDump(Address + 40 + SubIndex)
   Index = Index + 1
  Loop
  
  lstSpells.AddItem SpellName(Index)
  lstSpells.ItemData(lstSpells.ListCount - 1) = SpellCode(Index)
  
  Next SubIndex

  
  For Index = 0 To 6 Step 2
  
   If RomDump(Address + 33 + Index) > &H7F Then
     lstItems.AddItem ItemName(RomDump(Address + 33 + Index) - &H80) & " (Equipped)"
     lstItems.ItemData(lstItems.ListCount - 1) = RomDump(Address + 33 + Index)
   Else
     lstItems.AddItem ItemName(RomDump(Address + 33 + Index))
     lstItems.ItemData(lstItems.ListCount - 1) = RomDump(Address + 33 + Index)
   End If
   
  Next
  
  ''' Doing the resistance chart...
  
  Index = RomDump(Address + 27)
  
  If Index >= &H80 Then
   chk50Fire.Value = vbChecked
   Index = Index - &H80
  End If
  If Index >= &H40 Then
   chk25Fire.Value = vbChecked
   Index = Index - &H40
  End If
  If Index >= &H20 Then
   chk50Ice.Value = vbChecked
   Index = Index - &H20
  End If
  If Index >= &H10 Then
   chk25Ice.Value = vbChecked
   Index = Index - &H10
  End If
  If Index >= &H8 Then
   chk50Lightning.Value = vbChecked
   Index = Index - &H8
  End If
  If Index >= &H4 Then
   chk25Lightning.Value = vbChecked
   Index = Index - &H4
  End If
  If Index >= &H2 Then
   chk50Wind.Value = vbChecked
   Index = Index - &H2
  End If
  If Index >= &H1 Then
   chk25Wind.Value = vbChecked
   Index = Index - &H1
  End If
  
  Index = RomDump(Address + 26)
  
  
  If Index >= 128 Then
   Index = Index - 128
   chkMajor.Value = vbChecked
  End If
  If Index >= 64 Then
   Index = Index - 64
   chkMinor.Value = vbChecked
  End If
  
  SetCombo cboMoveType, CLng(RomDump(Address + 49))
  
  txtSpellpower.Text = RomDump(Address + 10)
  txtLevel.Text = RomDump(Address + 11)
  txtHPTimes.Text = RomDump(Address + 12)
  txtHP.Text = RomDump(Address + 13)
  txtMP.Text = RomDump(Address + 16)
  txtAtk.Text = RomDump(Address + 18)
  txtDef.Text = RomDump(Address + 20)
  txtAGI.Text = RomDump(Address + 22)
  txtMove.Text = RomDump(Address + 24)
  
  txtSpecial.Text = RomDump(Address + 30)
  
  ' Index = &H1F914 + cboName.ListIndex * 2
  Index = pBattleSprites + 270 + cboName.ListIndex * 2
  
  txtBattleModel.Text = RomDump(Index)
  txtBattlePalette.Text = RomDump(Index + 1)
    
    
  Index = ENEMYGOLD_ORIGINAL_OFFSET + cboName.ListIndex * 2
    
  txtGold.Text = CLng(RomDump(Index)) * 256 + RomDump(Index + 1)
    
  
  
  
  
    '' Critical chances and such crap
  
  Dim AtkChances As Byte
  
  AtkChances = RomDump(Address + 30)
  
  If CLng(AtkChances - 192) >= 0 Then
   optCounter(3).Value = True
  ElseIf CLng(AtkChances - 128) >= 0 Then
   optCounter(2).Value = True
  ElseIf CLng(AtkChances - 64) >= 0 Then
   optCounter(1).Value = True
  Else
   optCounter(0).Value = True
  End If
  
  If AtkChances - 128 >= 0 Then: AtkChances = AtkChances - 128
  If AtkChances - 64 >= 0 Then: AtkChances = AtkChances - 64
  
  If CLng(AtkChances - 48) >= 0 Then
   optDouble(3).Value = True
  ElseIf CLng(AtkChances - 32) >= 0 Then
   optDouble(2).Value = True
  ElseIf CLng(AtkChances - 16) >= 0 Then
   optDouble(1).Value = True
  Else
   optDouble(0).Value = True
  End If
  
  If AtkChances - 32 >= 0 Then: AtkChances = AtkChances - 32
  If AtkChances - 16 >= 0 Then: AtkChances = AtkChances - 16
  
  
  Dim Match As Boolean
  Dim Consider As Byte
  
  Match = False
  Consider = 15
  
  Do While Match = False
    If CLng(AtkChances - CLng(Consider)) >= 0 Then
     optCritical(Consider).Value = True
     Match = True
    End If
    
    If Consider >= 1 Then: Consider = Consider - 1
  Loop
  
  
  For AtkChances = 0 To 15
   optCritical(CInt(AtkChances)).Enabled = True
  Next AtkChances
  
  optCounter(0).Enabled = True
  optCounter(1).Enabled = True
  optCounter(2).Enabled = True
  optCounter(3).Enabled = True
  
  optDouble(0).Enabled = True
  optDouble(1).Enabled = True
  optDouble(2).Enabled = True
  optDouble(3).Enabled = True
  
  
  '' End special crap
  
  
  
  
  Address = cboName.ItemData(cboName.ListIndex)
  
  Loading = False
  
End Sub

Private Sub chkMajor_Click()
 txtNameData.Visible = False
 
 Call chkMinor_Click
End Sub

Private Sub chkMinor_Click()
 Dim Sum As Byte
 
 txtNameData.Visible = False
 
 If Loading = True Then
  Exit Sub
 End If
 
 If chkMinor.Value = vbChecked Then
  Sum = Sum + 64
 End If
 If chkMajor.Value = vbChecked Then
  Sum = Sum + 128
 End If
 
 RomDump(Address + 26) = Sum
End Sub

Private Sub cmdRename_Click()
 Dim Rename As New NameSpellHeroMonster
 Dim NameLocation As Integer

 Rename.Show
 
 NameLocation = 74 + CInt(cboName.ListIndex)
   
 Rename.cboName.ListIndex = NameLocation
 
End Sub

Private Sub Form_Load()
 
 Dim Index As Long
 Dim Count As Byte
 
 EnableTextboxes False, Me
  
 Index = pMonsterData     '&H1B1A66

 Do
 
  cboName.AddItem MonsterName(Count)
  cboName.ItemData(cboName.ListCount - 1) = Index
   
  Index = Index + 56
  Count = Count + 1
 
 Loop While Index <= pMonsterData + 5718      ' &H1B30BC
  
  
 For Index = 0 To UBound(SpellName())
   cboSpell.AddItem SpellName(Index)
   cboSpell.ItemData(Index) = SpellCode(Index)
 Next
 
 For Index = 0 To UBound(ItemName())
   cboItems.AddItem ItemName(Index)
   cboItems.ItemData(Index) = ItemCode(Index)
 Next
 
 For Index = 0 To UBound(MoveTypeName())
   cboMoveType.AddItem MoveTypeName(Index)
   cboMoveType.ItemData(Index) = MoveTypeCode(Index)
 Next
 
End Sub

Private Function Align(poo As Byte) As String

 If poo < 10 Then
  Align = "  "
 ElseIf poo < 100 Then
  Align = " "
 Else
  Align = ""
 End If

End Function

Private Sub lstItems_Click()
  
  txtNameData.Visible = False
  
  Loading = True
  
  If lstItems.ItemData(lstItems.ListIndex) > &H7F Then
   chkEquipped.Value = vbChecked
   SetCombo cboItems, CLng(lstItems.ItemData(lstItems.ListIndex) - &H80)
  Else
   chkEquipped.Value = vbUnchecked
   SetCombo cboItems, CLng(lstItems.ItemData(lstItems.ListIndex))
  End If
  
  Loading = False
  
End Sub

Private Sub lstSpells_Click()

 txtNameData.Visible = False

 Loading = True
 
 'Me.Caption = lstSpells.ItemData(lstSpells.ListIndex)
 
 SetCombo cboSpell, CLng(lstSpells.ItemData(lstSpells.ListIndex))
 
 Loading = False
End Sub

Private Sub cbospell_Click()
    
  txtNameData.Visible = False
  
  Address = cboName.ItemData(cboName.ListIndex)
  
  If lstSpells.ListIndex = -1 Or Loading = True Then
   Exit Sub
  End If
  
  lstSpells.List(lstSpells.ListIndex) = cboSpell.Text
  lstSpells.ItemData(lstSpells.ListIndex) = cboSpell.ItemData(cboSpell.ListIndex)
  
  RomDump(Address + 40 + lstSpells.ListIndex) = cboSpell.ItemData(cboSpell.ListIndex)
    
End Sub

Private Sub optCounter_Click(Index As Integer)
  If Loading = False Then: Call AssignDexterity
End Sub

Private Sub optCritical_Click(Index As Integer)
  If Loading = False Then: Call AssignDexterity
End Sub

Private Sub optDouble_Click(Index As Integer)
  If Loading = False Then: Call AssignDexterity
End Sub

Private Sub txtAGI_Change()
 Dim Result As Byte
 
 If ValidateInput(txtAGI, False, True) = True Then
  Result = txtAGI.Text
  RomDump(Address + 22) = Result
 End If
End Sub

Private Sub txtAGI_Click()
 txtNameData.Visible = False
End Sub

Private Sub txtATK_Change()
 Dim Result As Byte
 
 If ValidateInput(txtAtk, False, True) = True Then
  Result = txtAtk.Text
  RomDump(Address + 18) = Result
 End If
End Sub

Private Sub txtATK_Click()
 txtNameData.Visible = False
End Sub

Private Sub txtBattleModel_Change()
 Dim Result As Byte
 
 If ValidateInput(txtBattleModel, False, True) = True Then
  Result = txtBattleModel.Text
  ' RomDump(&H1F914 + cboName.ListIndex * 2) = Result
  RomDump(pBattleSprites + 270 + cboName.ListIndex * 2) = Result
 
 End If
End Sub

Private Sub txtBattleModel_Click()
 txtNameData.Visible = False
End Sub

Private Sub txtBattleModel_LostFocus()

 Dim Result As Byte
 
 If ValidateInput(txtBattleModel, False) = True Then
  Result = txtBattleModel.Text
  'RomDump(&H1F914 + cboName.ListIndex * 2) = Result
  RomDump(pBattleSprites + 270 + cboName.ListIndex * 2) = Result
 End If
  
End Sub

Private Sub txtBattlePalette_Change()
 Dim Result As Byte
 
 If ValidateInput(txtBattlePalette, False, True) = True Then
  Result = txtBattlePalette.Text
  'RomDump(&H1F914 + cboName.ListIndex * 2 + 1) = Result
  RomDump(pBattleSprites + 270 + cboName.ListIndex * 2 + 1) = Result
 End If
End Sub

Private Sub txtBattlePalette_Click()
 txtNameData.Visible = False
End Sub

Private Sub txtBattlePalette_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtBattlePalette, False) = True Then
  Result = txtBattlePalette.Text
  'RomDump(&H1F914 + cboName.ListIndex * 2 + 1) = Result
  RomDump(pBattleSprites + 270 + cboName.ListIndex * 2 + 1) = Result
 End If
End Sub

Private Sub txtDEF_Change()
 Dim Result As Byte
 
 If ValidateInput(txtDef, False, True) = True Then
  Result = txtDef.Text
  RomDump(Address + 20) = Result
 End If
End Sub

Private Sub txtDEF_Click()
 txtNameData.Visible = False
End Sub

Private Sub txtGold_Change()
 Dim Index As Long
 Dim Result As Byte
 Dim SubAddress As Long
 
 SubAddress = ENEMYGOLD_ORIGINAL_OFFSET + cboName.ListIndex * 2

 If ValidateInput(txtGold, True, True) = True Then
  
  Index = CLng(txtGold.Text)
  
  Result = CByte(Fix(Index / 256#))
  
  RomDump(SubAddress) = Result
  
  Result = CByte(Index - CLng(Result * 256#))
  
  RomDump(SubAddress + 1) = Result
  
 End If
End Sub

Private Sub txtGold_Click()
 txtNameData.Visible = False
End Sub

Private Sub txtGold_LostFocus()
 
 Dim Index As Long
 Dim Result As Byte
 Dim SubAddress As Long
 
 SubAddress = ENEMYGOLD_ORIGINAL_OFFSET + cboName.ListIndex * 2
 
 If ValidateInput(txtGold, True) = True Then
  
  Index = CLng(txtGold.Text)
  
  Result = CByte(Fix(Index / 256#))
  
  RomDump(SubAddress) = Result
  
  Result = CByte(Index - CLng(Result * 256#))
  
  RomDump(SubAddress + 1) = Result
  
 End If
  
End Sub

Private Sub txtHP_Change()
 Dim Result As Byte
 
 If ValidateInput(txtHP, False, True) = True Then
  Result = txtHP.Text
  RomDump(Address + 13) = Result
 End If
End Sub

Private Sub txtHP_Click()
 txtNameData.Visible = False
End Sub

Private Sub txtHPTimes_Change()
 Dim Result As Byte
 
 If ValidateInput(txtHPTimes, False, True) = True Then
  Result = txtHPTimes.Text
  RomDump(Address + 12) = Result
 End If
End Sub

Private Sub txtHPTimes_Click()
 txtNameData.Visible = False
End Sub

Private Sub txtLevel_Change()
 Dim Result As Byte
 
 If ValidateInput(txtLevel, False, True) = True Then
  Result = txtLevel.Text
  RomDump(Address + 11) = Result
 End If
End Sub

Private Sub txtLevel_Click()
 txtNameData.Visible = False
End Sub

Private Sub txtLevel_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtLevel, False) = True Then
  Result = txtLevel.Text
  RomDump(Address + 11) = Result
 End If
End Sub

Private Sub txtHp_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtHP, False) = True Then
  Result = txtHP.Text
  RomDump(Address + 13) = Result
 End If
End Sub

Private Sub txtHPTimes_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtHPTimes, False) = True Then
  Result = txtHPTimes.Text
  RomDump(Address + 12) = Result
 End If
End Sub

Private Sub txtMove_Change()
 Dim Result As Byte
 
 If ValidateInput(txtMove, False, True) = True Then
  Result = txtMove.Text
  RomDump(Address + 24) = Result
 End If
End Sub

Private Sub txtMove_Click()
 txtNameData.Visible = False
End Sub

Private Sub txtMP_Change()
 Dim Result As Byte
 
 If ValidateInput(txtMP, False, True) = True Then
  Result = txtMP.Text
  RomDump(Address + 16) = Result
 End If
End Sub

Private Sub txtMP_Click()
 txtNameData.Visible = False
End Sub

Private Sub txtMp_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtMP, False) = True Then
  Result = txtMP.Text
  RomDump(Address + 16) = Result
 End If
End Sub

Private Sub txtAtk_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtAtk, False) = True Then
  Result = txtAtk.Text
  RomDump(Address + 18) = Result
 End If
End Sub

Private Sub txtDef_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtDef, False) = True Then
  Result = txtDef.Text
  RomDump(Address + 20) = Result
 End If
End Sub

Private Sub txtAgi_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtAGI, False) = True Then
  Result = txtAGI.Text
  RomDump(Address + 22) = Result
 End If
End Sub

Private Sub txtMove_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtMove, False) = True Then
  Result = txtMove.Text
  RomDump(Address + 24) = Result
 End If
End Sub



Private Sub txtName_Change()

 Dim Index As Long
 Dim Result As String
 
 If Loading = False Then
 
 For Index = 1 To Len(txtName.Text)
  
  Result = Result & ChCode(Mid(txtName, Index, 1)) & vbNewLine
 
 Next Index
 
 Index = 0
 
 Do While Len(txtName.Text) + Index < txtName.MaxLength
 
  Result = Result & "32" & vbNewLine
  
  Index = Index + 1
 Loop
 
 txtNameData.Text = Left(Result, Len(Result) - 2)
 
 End If

End Sub

Private Sub txtName_Click()
 txtNameData.Visible = True
End Sub

Private Sub txtNameData_Change()
  On Error GoTo Skipo
  
  Dim Index As Long
  
  Dim NameArray() As String
  
  Dim Result As String
  
  NameArray = Split(txtNameData.Text, vbNewLine)

  For Index = 0 To UBound(NameArray())
   Result = Result & Chr(CLng(CByte(NameArray(Index))))
  Next Index
 
  For Index = 0 To UBound(NameArray())
  
   RomDump(NameAddress + Index) = CByte(NameArray(Index))
   
  Next Index
  
  
  Exit Sub
  
Skipo:
End Sub

Private Sub txtSpecial_Change()
 Dim Result As Byte
 
 If ValidateInput(txtSpecial, False, True) = True Then
  Result = txtSpecial.Text
  RomDump(Address + 30) = Result
 End If
End Sub

Private Sub txtSpecial_Click()
 txtNameData.Visible = False
End Sub

Private Sub txtSpecial_LostFocus()
 
 Dim Result As Byte
 
 If ValidateInput(txtSpecial, False) = True Then
  Result = txtSpecial.Text
  RomDump(Address + 30) = Result
 End If
 
End Sub

Private Sub txtNameData_KeyDown(KeyCode As Integer, Shift As Integer)
 
 DoEvents
 
 If KeyCode = 13 Then
  txtNameData.Text = Left(txtNameData.Text, txtNameData.SelStart - 2) & Right(txtNameData.Text, Len(txtNameData.Text) - txtNameData.SelStart)
 End If
End Sub

Private Function GetNameAddress(ByVal iName As Long) As Long

 Dim Index As Long
 Dim Count As Long
 
 Count = ENEMYNAMES_ORIGINAL_OFFSET

 Do While Index < iName
  Count = Count + RomDump(Count) + 1
  
  Index = Index + 1
 Loop
 
 GetNameAddress = Count + 1
 
End Function

Private Function AssignDexterity()

 Dim Index As Integer
 Dim Value As Byte
   
  Value = 0
  
  Select Case True
  Case optCounter(3).Value
   Value = Value + 192
  Case optCounter(2).Value
   Value = Value + 128
  Case optCounter(1).Value
   Value = Value + 64
  Case optCounter(0).Value
   Value = Value + 0
  End Select
  
  Select Case True
  Case optDouble(3).Value
   Value = Value + 48
  Case optDouble(2).Value
   Value = Value + 32
  Case optDouble(1).Value
   Value = Value + 16
  Case optDouble(0).Value
   Value = Value + 0
  End Select
  
  For Index = 0 To 15
   If optCritical(Index).Value = True Then: Value = Value + CByte(Index)
  Next Index


  RomDump(Address + 30) = Value

End Function


Private Sub txtSpellpower_Change()
 Dim Result As Byte
 
 If ValidateInput(txtSpellpower, False, True) = True Then
  Result = txtSpellpower.Text
  RomDump(Address + 10) = Result
 End If
End Sub

Private Sub txtSpellpower_Click()
 txtNameData.Visible = False
End Sub
