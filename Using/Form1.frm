VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBScript.Regexp - PCRE2 ProxyWrapper"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   12450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTab 
      Caption         =   "Replace"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CheckBox chkTab 
      Caption         =   "Match"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CheckBox chkTab 
      Caption         =   "Compile"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3600
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox txtError 
      Height          =   1095
      Left            =   960
      TabIndex        =   36
      Top             =   7320
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1931
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame FraEngine 
      Caption         =   "Engine"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   10440
      TabIndex        =   18
      Top             =   240
      Width           =   1935
      Begin VB.CheckBox chkSuppressErrors 
         Caption         =   "Suppress errors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.OptionButton OptEngineBoth 
         Caption         =   "Both"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton OptEngineVB 
         Caption         =   "VBScript.Regexp"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton OptEnginePCRE 
         Caption         =   "PCRE2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame FraOptionsCommon 
      Caption         =   "Common options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      TabIndex        =   13
      Top             =   2400
      Width           =   4215
      Begin VB.CheckBox chkIgnoreCase 
         Caption         =   "Ignore Case"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   17
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkMultiline 
         Caption         =   "Multiline"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   16
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkGlobal 
         Caption         =   "Global"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8160
      TabIndex        =   10
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8160
      TabIndex        =   9
      Top             =   960
      Width           =   2055
   End
   Begin RichTextLib.RichTextBox txtPattern 
      Height          =   615
      Left            =   960
      TabIndex        =   8
      Top             =   1560
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":007D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtSource 
      Height          =   1215
      Left            =   960
      TabIndex        =   7
      Top             =   240
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   2143
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0154
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtPCRE 
      Height          =   3615
      Left            =   4080
      TabIndex        =   4
      Top             =   3600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6376
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"Form1.frx":0225
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtVB 
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6376
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"Form1.frx":02A0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Execute"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8160
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin RichTextLib.RichTextBox txtReplace 
      Height          =   615
      Left            =   960
      TabIndex        =   11
      Top             =   2280
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":031B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame FraOptionsMatch 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   12600
      TabIndex        =   39
      Top             =   3840
      Width           =   4215
      Begin VB.CheckBox chkMatchAnchored 
         Caption         =   "Anchored"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   360
         Width           =   3735
      End
      Begin VB.CheckBox chkMatchMatchedEventEnabled 
         Caption         =   "Matched Event Enabled"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   720
         Width           =   3735
      End
      Begin VB.CheckBox chkMatchMatchIfEmpty 
         Caption         =   "Match If Empty"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   1080
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkMatchMatchIfEmptyAtStart 
         Caption         =   "Match If Empty At Start"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   1440
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkMatchPerformUtfCheck 
         Caption         =   "Perform Utf Check"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   1800
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkMatchSubjectIsBeginningOfLine 
         Caption         =   "Subject Is Beginning Of Line"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   2160
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkMatchSubjectIsEndOfLine 
         Caption         =   "Subject Is End Of Line"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   2520
         Value           =   1  'Checked
         Width           =   3735
      End
   End
   Begin VB.Frame FraOptionsReplace 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   16920
      TabIndex        =   47
      Top             =   3840
      Width           =   4215
      Begin VB.CheckBox chkReplaceTreatUnknownCapturingGroupsAsEmptyStrings 
         Caption         =   "Treat Unknown Capturing Groups As Empty Strings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   56
         Top             =   3240
         Width           =   3735
      End
      Begin VB.CheckBox chkReplaceErrorOnUnknownCapturingGroups 
         Caption         =   "Error On Unknown Capturing Groups"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   2880
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkReplaceSubjectIsEndOfLine 
         Caption         =   "Subject Is End Of Line"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   2520
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkReplaceSubjectIsBeginningOfLine 
         Caption         =   "Subject Is Beginning Of Line"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   2160
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkReplacePerformUtfCheck 
         Caption         =   "Perform Utf Check"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   1800
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkReplaceMatchIfEmptyAtStart 
         Caption         =   "Match If Empty At Start"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   1440
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkReplaceMatchIfEmpty 
         Caption         =   "Match If Empty"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   1080
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkReplaceExtendedReplacement 
         Caption         =   "Extended Replacement"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   720
         Width           =   3735
      End
      Begin VB.CheckBox chkReplaceAnchored 
         Caption         =   "Anchored"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame FraOptionsCompile 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   8160
      TabIndex        =   14
      Top             =   3840
      Width           =   4215
      Begin VB.CheckBox chkIgnorePatternWhitspaceAndComments 
         Caption         =   "Ignore Pattern Whitspace And Comments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   3960
         Width           =   3855
      End
      Begin VB.CheckBox chkGreedy 
         Caption         =   "Greedy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   3600
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkDotMatchesAllCharacters 
         Caption         =   "Dot Matches All Characters"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   3240
         Width           =   3735
      End
      Begin VB.CheckBox chkDollarMatchesEndOfStringOnly 
         Caption         =   "Dollar Matches End Of String Only"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   2880
         Width           =   3735
      End
      Begin VB.CheckBox chkUtf 
         Caption         =   "Utf"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   2520
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkCheckUtfValidity 
         Caption         =   "Check Utf Validity"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   2160
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkAnchored 
         Caption         =   "Anchored"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1800
         Width           =   3735
      End
      Begin VB.CheckBox chkAlternateCircumflexHandling 
         Caption         =   "Alternate Circumflex Handling"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1440
         Width           =   3735
      End
      Begin VB.CheckBox chkAlternateBsuxHandling 
         Caption         =   "Alternate Bsux Handling"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkAllowEmptyClass 
         Caption         =   "Allow Empty Class"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   3735
      End
      Begin VB.CheckBox chkAllowDuplicateSubpatternNames 
         Caption         =   "Allow Duplicate Subpattern Names"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3450
      TabIndex        =   59
      Top             =   3300
      Width           =   1245
   End
   Begin VB.Label lblPCRE3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PCRE2 options:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8280
      TabIndex        =   38
      Top             =   3240
      Width           =   1245
   End
   Begin VB.Label lblErrors 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Errors:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   34
      Top             =   7320
      Width           =   525
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   8160
      X2              =   0
      Y1              =   7245
      Y2              =   7245
   End
   Begin VB.Label lblResult 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Results"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3720
      TabIndex        =   33
      Top             =   3000
      Width           =   720
   End
   Begin VB.Label lblReplace 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Replace (optional)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   975
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPCRE2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PCRE2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5520
      TabIndex        =   6
      Top             =   3240
      Width           =   555
   End
   Begin VB.Label lblVBScriptRegexp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VBScript.Regexp"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1200
      TabIndex        =   5
      Top             =   3240
      Width           =   1410
   End
   Begin VB.Label lblSourceText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pattern"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   795
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Source Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   795
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' -----------------------------------------
'
'          ProxyWrapper object model
'
' -----------------------------------------
'
' It's a full mimic of VBScript.Regexp,
' + .UsePcre property => to switch beetween engines;
' + .PCRE2 property => for accessing directly to PCRE2 wrapper object model.
'
' IRegExp ->
'   .Global      as boolean
'   .IgnoreCase  as boolean
'   .Multiline   as boolean
'   .Pattern     as string
'   .Replace     (sourceString As String, replaceVar As Variant) As String
'   .Test        (sourceString As String) As Boolean
'   .Execute     (sourceString As String) As IRegExpMatchCollection
'   .UsePcre     as boolean
'   .PCRE2       as IPcre
'
' IRegExpMatchCollection ->
'   .Count       as long
'   .Item        (Index As Long) As IRegExpMatch
'
' IRegExpMatch ->
'   .FirstIndex  as long
'   .Length      as long
'   .Value       as string
'   SubMatches() As IRegExpSubMatches
'
' IRegExpSubMatches ->
'   .Count       as long
'   .Item        (Index As Long) As String

'---------------------------------------------------
' For RichTextBox unicode support
'
Private Type GETTEXTLENGTHEX
    flags       As Long
    codepage    As Long
End Type

Private Type GETTEXTEX
    cb              As Long
    flags           As Long
    codepage        As Long
    lpDefaultChar   As Long
    lpUsedDefChar   As Long
End Type

Private Type SETTEXTEX
    flags       As Long
    codepage    As Long
End Type

Private Declare Function SendMessageW Lib "user32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const CP_UNICODE            As Long = 1200&
Private Const GT_USECRLF            As Long = 1&
Private Const GTL_USECRLF           As Long = 1&
Private Const GTL_PRECISE           As Long = 2&
Private Const GTL_NUMCHARS          As Long = 8&
Private Const ST_UNICODE            As Long = 8&

Private Const WM_USER               As Long = &H400&
Private Const EM_GETTEXTEX          As Long = WM_USER + 94
Private Const EM_GETTEXTLENGTHEX    As Long = WM_USER + 95
Private Const EM_SETTEXTEX          As Long = 1121&
'---------------------------------------------------


' -----------------------------------------
'
'          Examples of using
'
' -----------------------------------------

Dim mo_Regexp As IRegExp    'For GUI (main form)


Private Sub Form_Load()

    'For GUI
    Set mo_Regexp = New cRegExp
    mo_Regexp.UsePcre = True
    ApplySettings
    
    
    'Stand-alone examples
    
    ' 1. "Test" method - Verify that the pattern matches.
    TestTest "https://mail.ya.ru", "^(http(s)?://|www(2)?\.)?([^/]*\.)?ya.ru(/|$|\?)", bUsePcre:=False  'with VBScript.Regexp
    TestTest "https://mail.ya.ru", "^(http(s)?://|www(2)?\.)?([^/]*\.)?ya.ru(/|$|\?)", bUsePcre:=True   'with PCRE2
    
    ' 2. "Replace" method - Replace matches of pattern in original string by specified substring.
    ' Hack (^_^): Multiply the number of dollars per 1000
    TestReplace "3 dol. 5 cents.", "(\d)+\s+(dol|dollars|d)s?\.?\s+(.*)", "$1.000 dollars $3", bUsePcre:=False     'with VBScript.Regexp
    TestReplace "3 dol. 5 cents.", "(\d)+\s+(dol|dollars|d)s?\.?\s+(.*)", "$1.000 dollars $3", bUsePcre:=True      'with PCRE2
    
    ' 3. "Execute" method - Returns the collection of found substrings by pattern.
    TestExecute "File1.zip.exe" & vbCrLf & "File2.com" & vbCrLf & "File 3", "[\w ]+(\.\S+?)*$", bUsePcre:=False     'with VBScript.Regexp
    TestExecute "File1.zip.exe" & vbCrLf & "File2.com" & vbCrLf & "File 3", "[\w ]+(\.\S+?)*$", bUsePcre:=True      'with PCRE2
    
    'Unload me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set mo_Regexp = Nothing 'for GUI
End Sub


'
' "Test" method
'
Sub TestTest(p_SubjectText As String, p_Regex As String, bUsePcre As Boolean)

    Debug.Print vbCrLf & Space(15) & "Test method - " & IIf(bUsePcre, "PCRE2", "VBScript")
    
    ' declaration
    Dim lo_RegEx As IRegExp
    ' creating an instance
    Set lo_RegEx = New cRegExp
    ' set pattern
    lo_RegEx.Pattern = p_Regex
    ' choose engine
    lo_RegEx.UsePcre = bUsePcre
    ' run "Test" method
    Debug.Print lo_RegEx.Test(p_SubjectText)
    ' destroy an instance of the class
    Set lo_RegEx = Nothing
End Sub


'
' "Replace" method
'
Sub TestReplace(p_SubjectText As String, p_Regex As String, p_ReplaceText As String, bUsePcre As Boolean)

    Debug.Print vbCrLf & Space(15) & "Replace method - " & IIf(bUsePcre, "PCRE2", "VBScript")

    Dim lo_RegEx As IRegExp
    Set lo_RegEx = New cRegExp
    
    lo_RegEx.Pattern = p_Regex
    lo_RegEx.UsePcre = bUsePcre

    Debug.Print lo_RegEx.Replace(p_SubjectText, p_ReplaceText)

    Set lo_RegEx = Nothing
End Sub


'
' "Execute" method + iterating submatches
'
Sub TestExecute(p_SubjectText As String, p_Regex As String, bUsePcre As Boolean)
   
   Debug.Print vbCrLf & Space(15) & "Execute Method - " & IIf(bUsePcre, "PCRE2", "VBScript")
   
'   'declarations
'   Dim lo_RegEx         As IRegExp
'   Dim lo_Matches       As Object
'   Dim lo_Match         As Object
'   Dim lo_Submatches    As Object
'   Dim lo_SubMatch      As Variant

'   'alternate declarations (just to support IntelliSense)
   Dim lo_RegEx         As IRegExp                  'VBScript_RegExp_55.RegExp
   Dim lo_Matches       As IRegExpMatchCollection   'VBScript_RegExp_55.MatchCollection
   Dim lo_Match         As IRegExpMatch             'VBScript_RegExp_55.Match
   Dim lo_Submatches    As IRegExpSubMatches        'VBScript_RegExp_55.SubMatches
   Dim lo_SubMatch      As Variant
   
   Dim l_SubjectText As String
   Dim l_Regex As String
   
   Dim ii As Long
   Dim jj As Long
   
   l_SubjectText = "File1.zip.exe" & vbCrLf & "File2.com" & vbCrLf & "File 3"
   l_Regex = "[\w ]+(\.\S+?)*$"
   
   ' creating an instance
   Set lo_RegEx = New cRegExp
   
   ' settings
   With lo_RegEx
      .IgnoreCase = True
      .Global = True
      .MultiLine = True
      .Pattern = l_Regex
      .UsePcre = bUsePcre ' set whether we want to use PCRE2 or VBScript.Regexp version
   End With
   
   'run "Execute" method
   Set lo_Matches = lo_RegEx.Execute(l_SubjectText)
   
   Debug.Print "Match Count: " & lo_Matches.Count

   'iterating items
   For Each lo_Match In lo_Matches
    
      'Debug.Print "Print value by index of collection: " & lo_Matches(0).Value
    
      Set lo_Submatches = lo_Match.SubMatches
    
      ii = ii + 1
      Debug.Print "Match #" & ii & ": " & lo_Match.Value
      'Debug.Print "Sub Match Count: " & lo_Submatches.Count
      
      'iterating submatches
      jj = 0
      For Each lo_SubMatch In lo_Submatches
        jj = jj + 1
        Debug.Print Space$(25) & "SubMatch # " & jj & ": " & lo_SubMatch
      Next
      
'      'alternate way
'      For jj = 0 To lo_Submatches.Count - 1
'        Debug.Print "SubMatch # " & jj + 1 & ": " & lo_Submatches.Item(jj)  'alternate 1
'        Debug.Print "SubMatch # " & jj + 1 & ": " & lo_Submatches(jj)       'alternate 2
'      Next
   Next
   
   ' destroy an instance of the class
   Set lo_RegEx = Nothing
End Sub


' -------------------------------------------------------------
'
'                        GUI (main form)
'
' -------------------------------------------------------------

Private Sub ApplySettings()
    'Apply options, checked on the main form
    
    With mo_Regexp.PCRE2.Options
      With .Compile
        .AllowDuplicateSubpatternNames = chkAllowDuplicateSubpatternNames.Value
        .AllowEmptyClass = chkAllowEmptyClass.Value
        .AlternateBsuxHandling = chkAlternateBsuxHandling.Value
        .AlternateCircumflexHandling = chkAlternateCircumflexHandling.Value
        .Anchored = chkAnchored.Value
        .CheckUtfValidity = chkCheckUtfValidity.Value
        .Utf = chkUtf.Value
        .DollarMatchesEndOfStringOnly = chkDollarMatchesEndOfStringOnly.Value
        .DotMatchesAllCharacters = chkDotMatchesAllCharacters.Value
        .Greedy = chkGreedy.Value
        .IgnorePatternWhitspaceAndComments = chkIgnorePatternWhitspaceAndComments.Value
      End With
      With .Match
        .Anchored = chkMatchAnchored.Value
        .MatchedEventEnabled = chkMatchMatchedEventEnabled.Value
        .MatchIfEmpty = chkMatchMatchIfEmpty.Value
        .MatchIfEmptyAtStart = chkMatchMatchIfEmptyAtStart.Value
        .PerformUtfCheck = chkMatchPerformUtfCheck.Value
        .SubjectIsBeginningOfLine = chkMatchSubjectIsBeginningOfLine.Value
        .SubjectIsEndOfLine = chkMatchSubjectIsEndOfLine.Value
      End With
      With .Replace
        .Anchored = chkReplaceAnchored.Value
        .ErrorOnUnknownCapturingGroups = chkReplaceErrorOnUnknownCapturingGroups.Value
        .ExtendedReplacement = chkReplaceExtendedReplacement.Value
        .MatchIfEmpty = chkReplaceMatchIfEmpty.Value
        .MatchIfEmptyAtStart = chkReplaceMatchIfEmptyAtStart.Value
        .PerformUtfCheck = chkReplacePerformUtfCheck.Value
        .SubjectIsBeginningOfLine = chkReplaceSubjectIsBeginningOfLine.Value
        .SubjectIsEndOfLine = chkReplaceSubjectIsEndOfLine.Value
        .TreatUnknownCapturingGroupsAsEmptyStrings = chkReplaceTreatUnknownCapturingGroupsAsEmptyStrings.Value
      End With
    End With
    
    'common options
    mo_Regexp.Global = chkGlobal.Value
    mo_Regexp.IgnoreCase = chkIgnoreCase.Value
    mo_Regexp.MultiLine = chkMultiline.Value
    mo_Regexp.Pattern = RTBReadUnicode(txtPattern)
End Sub

Private Sub chkTab_Click(Index As Integer)
    Static bClicked As Boolean
    If Not bClicked Then
        bClicked = True
        chkTab(0).Value = 0
        chkTab(1).Value = 0
        chkTab(2).Value = 0
        chkTab(Index).Value = 1
        FraOptionsCompile.Visible = False
        FraOptionsMatch.Visible = False
        FraOptionsReplace.Visible = False
        If Index = 0 Then FraOptionsCompile.Visible = True
        If Index = 1 Then FraOptionsMatch.Left = FraOptionsCompile.Left: FraOptionsMatch.Visible = True
        If Index = 2 Then FraOptionsReplace.Left = FraOptionsCompile.Left: FraOptionsReplace.Visible = True
        txtSource.SetFocus
        bClicked = False
    End If
End Sub

Private Sub ClearFields()
    txtVB.Text = ""
    txtPCRE.Text = ""
    txtError.Text = ""
    lblStatus.Caption = ""
End Sub

Private Sub cmdExecute_Click()
    If chkSuppressErrors.Value = 1 Then On Error GoTo ErrorHandler
    Dim l_EngineName As String
    Dim i&
    
    ClearFields
    ApplySettings
    
    Dim lo_Matches       As Object
    Dim lo_Match         As Object
    Dim lo_Submatches    As Object
    Dim lo_SubMatch      As Variant
    Dim ii&, jj&
    
    For i = 1 To 2
    
        If i = 1 Then
            l_EngineName = "VBScript.Regexp"
            mo_Regexp.UsePcre = False
        Else
            l_EngineName = "PCRE2"
            mo_Regexp.UsePcre = True
        End If
        
        ii = 0
        jj = 0
        
        If chkSuppressErrors.Value = 1 Then On Error Resume Next
        Err.Clear
        Set lo_Matches = mo_Regexp.Execute(RTBReadUnicode(txtSource))   'run "Execute" method
        
        If Err.Number <> 0 Then
            
            PrintText l_EngineName, "Error " & Err.Number & " - " & Err.Description
        
            If chkSuppressErrors.Value = 1 Then On Error GoTo ErrorHandler Else On Error GoTo 0
        Else
            If chkSuppressErrors.Value = 1 Then On Error GoTo ErrorHandler Else On Error GoTo 0
            
            PrintText l_EngineName, "Match Count: " & lo_Matches.Count
            PrintText l_EngineName, ""
    
            For Each lo_Match In lo_Matches
        
                Set lo_Submatches = lo_Match.SubMatches
        
                ii = ii + 1
                PrintText l_EngineName, "#" & ii & ": " & lo_Match.Value
    
                For Each lo_SubMatch In lo_Submatches
                    jj = jj + 1
                    PrintText l_EngineName, Space(30) & "Sub.#" & jj & ": " & lo_SubMatch
                Next
            Next
        End If
    
    Next
    
    CheckDifference
    Exit Sub
ErrorHandler:
    txtError.Text = l_EngineName & ": " & Err.Description
End Sub

Private Sub cmdTest_Click()
    If chkSuppressErrors.Value = 1 Then On Error GoTo ErrorHandler
    Dim l_EngineName As String
    Dim bResult As Boolean
    
    ClearFields
    ApplySettings
    
    If chkSuppressErrors.Value = 1 Then On Error Resume Next
    
    l_EngineName = "VBScript.Regexp"
    mo_Regexp.UsePcre = False
    Err.Clear
    bResult = mo_Regexp.Test(RTBReadUnicode(txtSource))
    
    If Err.Number <> 0 Then
        PrintText l_EngineName, "Error " & Err.Number & " - " & Err.Description
    Else
        PrintText l_EngineName, CStr(bResult) 'run "Test" method
    End If
    
    l_EngineName = "PCRE2"
    mo_Regexp.UsePcre = True
    Err.Clear
    bResult = mo_Regexp.Test(RTBReadUnicode(txtSource))
    
    If Err.Number <> 0 Then
        PrintText l_EngineName, "Error " & Err.Number & " - " & Err.Description
    Else
        PrintText l_EngineName, CStr(bResult) 'run "Test" method
    End If
    
    If chkSuppressErrors.Value = 1 Then On Error GoTo ErrorHandler Else On Error GoTo 0
    
    CheckDifference
    Exit Sub
ErrorHandler:
    txtError.Text = l_EngineName & ": " & Err.Description
End Sub

Private Sub cmdReplace_Click()
    If chkSuppressErrors.Value = 1 Then On Error GoTo ErrorHandler
    Dim l_EngineName As String
    Dim sResult As String
    
    ClearFields
    ApplySettings
    
    If chkSuppressErrors.Value = 1 Then On Error Resume Next
    
    l_EngineName = "VBScript.Regexp"
    mo_Regexp.UsePcre = False
    Err.Clear
    sResult = mo_Regexp.Replace(RTBReadUnicode(txtSource), RTBReadUnicode(txtReplace))
    
    If Err.Number <> 0 Then
        PrintText l_EngineName, "Error " & Err.Number & " - " & Err.Description
    Else
        PrintText l_EngineName, sResult  'run "Replace" method
    End If
    
    l_EngineName = "PCRE2"
    mo_Regexp.UsePcre = True
    Err.Clear
    
    If Err.Number <> 0 Then
        PrintText l_EngineName, "Error " & Err.Number & " - " & Err.Description
    Else
        PrintText l_EngineName, sResult  'run "Replace" method
    End If
    
    If chkSuppressErrors.Value = 1 Then On Error GoTo ErrorHandler Else On Error GoTo 0
    
    CheckDifference
    Exit Sub
ErrorHandler:
    txtError.Text = l_EngineName & ": " & Err.Description
End Sub

Private Sub CheckDifference()
    If OptEngineBoth.Value Then
        lblStatus.Caption = IIf(RTBReadUnicode(txtVB) = RTBReadUnicode(txtPCRE), "Same", "Different")
    Else
        lblStatus.Caption = ""
    End If
End Sub

Private Sub PrintText(p_EngineName As String, p_Text As String)
    Dim l_CurText As String
    Dim lo_RTB As RichTextBox
    
    If p_EngineName = "VBScript.Regexp" Then Set lo_RTB = txtVB
    If p_EngineName = "PCRE2" Then Set lo_RTB = txtPCRE
    
    l_CurText = RTBReadUnicode(lo_RTB)
    RTBWriteUnicode lo_RTB, l_CurText & IIf(Len(l_CurText) = 0, "", vbCrLf) & p_Text
End Sub

Public Function RTBReadUnicode(ByVal RTB As RichTextLib.RichTextBox) As String
    'Reads Text from RichTextBox as Unicode text
    Dim gtlUnicode As GETTEXTLENGTHEX
    Dim gtUnicode As GETTEXTEX
    Dim lngChars As Long

    With gtlUnicode
        .flags = GTL_USECRLF Or GTL_PRECISE Or GTL_NUMCHARS
        .codepage = CP_UNICODE
    End With
    lngChars = SendMessageW(RTB.hWnd, EM_GETTEXTLENGTHEX, VarPtr(gtlUnicode), ByVal 0)

    With gtUnicode
        .cb = (lngChars + 1) * 2
        .flags = GT_USECRLF
        .codepage = CP_UNICODE
    End With
    RTBReadUnicode = String$(lngChars, 0)
    SendMessageW RTB.hWnd, EM_GETTEXTEX, VarPtr(gtUnicode), ByVal StrPtr(RTBReadUnicode)
End Function

Public Function RTBWriteUnicode(ByVal RTB As RichTextLib.RichTextBox, p_Text As String) As String
    'Write Text to RichTextBox as Unicode
    Dim stUnicode As SETTEXTEX

    With stUnicode
        .flags = ST_UNICODE
        .codepage = CP_UNICODE
    End With
    SendMessageW RTB.hWnd, EM_SETTEXTEX, VarPtr(stUnicode), ByVal StrPtr(p_Text)
End Function
