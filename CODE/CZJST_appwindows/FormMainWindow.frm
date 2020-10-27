VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form FormMainWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fight Landlord Card Game Assistant　v1.01　by Sam Toki"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   795
   ClientWidth     =   15030
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "FormMainWindow.frx":0000
   LinkTopic       =   "FormMainWindow"
   MaxButton       =   0   'False
   MouseIcon       =   "FormMainWindow.frx":000C
   MousePointer    =   99  'Custom
   ScaleHeight     =   7815
   ScaleWidth      =   15030
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame FrameDoubler 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Doubler"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5790
      Left            =   210
      TabIndex        =   7
      Top             =   1785
      Width           =   4950
      Begin VB.CommandButton CmdDoublerX6 
         Caption         =   "x6"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   3155
         MouseIcon       =   "FormMainWindow.frx":015E
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   4100
         Width           =   1380
      End
      Begin VB.CommandButton CmdDoublerUndo 
         Caption         =   "←"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   1785
         MouseIcon       =   "FormMainWindow.frx":02B0
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   4100
         Width           =   1380
      End
      Begin VB.CommandButton CmdDoublerReset 
         Caption         =   "＊"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   420
         MouseIcon       =   "FormMainWindow.frx":0402
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   4100
         Width           =   1380
      End
      Begin VB.CommandButton CmdDoublerX5 
         Caption         =   "x5"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   3155
         MouseIcon       =   "FormMainWindow.frx":0554
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   2835
         Width           =   1380
      End
      Begin VB.CommandButton CmdDoublerX3 
         Caption         =   "x3"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   1785
         MouseIcon       =   "FormMainWindow.frx":06A6
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   2835
         Width           =   1380
      End
      Begin VB.TextBox TextboxDoublerNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   1725
         Left            =   420
         Locked          =   -1  'True
         MouseIcon       =   "FormMainWindow.frx":07F8
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Text            =   "1"
         Top             =   630
         Width           =   4095
      End
      Begin VB.CommandButton CmdDoublerX2 
         Caption         =   "x2"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   420
         MouseIcon       =   "FormMainWindow.frx":094A
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   2835
         Width           =   1380
      End
   End
   Begin VB.Frame FrameDice 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Dice"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7365
      Left            =   5460
      TabIndex        =   0
      Top             =   210
      Width           =   9360
      Begin VB.CommandButton CmdDiceReset 
         Caption         =   "＊"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   420
         MouseIcon       =   "FormMainWindow.frx":0A9C
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   5670
         Width           =   1380
      End
      Begin VB.CommandButton CmdDiceRoll 
         Caption         =   "START"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   2310
         MouseIcon       =   "FormMainWindow.frx":0BEE
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   5670
         Width           =   6630
      End
      Begin VB.Timer TimerDice 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   8715
         Top             =   6720
      End
      Begin VB.TextBox TextboxDiceNumber2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   156
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   3645
         Left            =   4880
         Locked          =   -1  'True
         MouseIcon       =   "FormMainWindow.frx":0D40
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Text            =   "?"
         Top             =   1575
         Width           =   4050
      End
      Begin VB.TextBox TextboxDiceNumber1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   156
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   3645
         Left            =   420
         Locked          =   -1  'True
         MouseIcon       =   "FormMainWindow.frx":0E92
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Text            =   "?"
         Top             =   1575
         Width           =   4050
      End
      Begin VB.Label LabelDiceNumber2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   795
         Left            =   4880
         TabIndex        =   5
         Top             =   525
         Width           =   4035
      End
      Begin VB.Label LabelDiceNumber1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Dice"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   795
         Left            =   420
         TabIndex        =   3
         Top             =   525
         Width           =   4035
      End
   End
   Begin VB.Timer TimerClock 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Label LabelClock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   56.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1200
      Left            =   210
      TabIndex        =   15
      Top             =   315
      Width           =   4965
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   420
      Left            =   0
      TabIndex        =   16
      Top             =   525
      Visible         =   0   'False
      Width           =   450
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   794
      _cy             =   741
   End
   Begin VB.Menu MenuDoubler 
      Caption         =   "D&oubler"
      Begin VB.Menu MenuDoublerX2 
         Caption         =   "x2"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu MenuDoublerX3 
         Caption         =   "x3"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu MenuDoublerX5 
         Caption         =   "x5"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu MenuDoublerX6 
         Caption         =   "x6"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu MenuDoubler1_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuDoublerUndo 
         Caption         =   "←　Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu MenuDoublerReset 
         Caption         =   "＊　Reset"
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu MenuDice 
      Caption         =   "D&ice"
      Begin VB.Menu MenuDiceRoll 
         Caption         =   "START"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MenuDice1_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuDiceReset 
         Caption         =   "＊　Reset"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu Menu1_ 
      Caption         =   "　|　"
      Enabled         =   0   'False
   End
   Begin VB.Menu MenuSoundSwitch 
      Caption         =   "Soun&d ON"
   End
   Begin VB.Menu MenuAbout 
      Caption         =   "&About"
      Begin VB.Menu MenuAboutName 
         Caption         =   "Fight Landlord Card Game Assistant"
      End
      Begin VB.Menu MenuAboutVersion 
         Caption         =   "v1.01 Release Version　|　for Windows 7,8,10　|　Multilingual"
      End
      Begin VB.Menu MenuAboutDate 
         Caption         =   "Last compiled on Thu, Sep 24, 2020"
      End
      Begin VB.Menu MenuAboutFirst 
         Caption         =   "First version built on Sun, Jul 14, 2019"
      End
      Begin VB.Menu MenuAbout1_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAboutAuthor 
         Caption         =   "Author: Sam Toki"
      End
      Begin VB.Menu MenuAboutOrganization 
         Caption         =   "Organization: SAM TOKI STUDIO"
      End
      Begin VB.Menu MenuAboutFrom 
         Caption         =   "From: Xidian University, China"
      End
      Begin VB.Menu MenuAboutContact 
         Caption         =   "Contact: SamToki@outlook.com"
      End
      Begin VB.Menu MenuAbout2_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAboutCopyright 
         Caption         =   "TM ＆ (C) 2015-2020 SAM TOKI STUDIO. All rights reserved."
      End
      Begin VB.Menu MenuAboutTrademark 
         Caption         =   "SAM TOKI STUDIO is a trademark of CZJ Software Technologies (CZJST) Inc. in the P.R.C and other countries."
      End
      Begin VB.Menu MenuAbout3_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAboutCommercial 
         Caption         =   "Commercial use of this software is strictly prohibited."
      End
   End
   Begin VB.Menu Menu2_ 
      Caption         =   "　|　"
      Enabled         =   0   'False
   End
   Begin VB.Menu MenuLanguage 
      Caption         =   "Ａ字あ (&L)"
      Begin VB.Menu MenuLanguageENG 
         Caption         =   "English (United States)"
         Checked         =   -1  'True
         Shortcut        =   +{F1}
      End
      Begin VB.Menu MenuLanguageCHS 
         Caption         =   "中文（简体）"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu MenuLanguageCHT 
         Caption         =   "中文（繁體）"
         Enabled         =   0   'False
         Shortcut        =   +{F3}
      End
      Begin VB.Menu MenuLanguageJPN 
         Caption         =   "日本語"
         Shortcut        =   +{F4}
      End
   End
   Begin VB.Menu Menu3_ 
      Caption         =   "　|　"
      Enabled         =   0   'False
   End
   Begin VB.Menu MenuEXIT 
      Caption         =   "E&XIT"
   End
End
Attribute VB_Name = "FormMainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[] DIM []

Option Explicit

'DIM Features...
Public dicestatus As Single
Public dicenumber1 As Single
Public dicenumber2 As Single
Public dicetemp As Single
Public dicecounter As Single
Public doublernumber As Long
Public doublernumberprev As Long

'DIM Preferences...
Public setlanguage As String
Public soundswitch As Boolean
Public windowanimationswitch As Boolean

'DIM Clock...
Public clockhour As Single
Public clockmin As Single
Public clocksec As Single

'DIM Dialogue...
Public answer

'================================================================================

'================================================================================

'[] LOAD []

    Public Sub Form_Load()
        'FIRST GENERAL RESET

        dicestatus = 0
        dicenumber1 = 0
        dicenumber2 = 0
        dicetemp = 0
        dicecounter = 0
        doublernumber = 1
        doublernumberprev = 1
        setlanguage = "ENG"
        soundswitch = True
        windowanimationswitch = True

        TextboxDiceNumber1.Text = "?"
        TextboxDiceNumber2.Text = "?"
        TextboxDoublerNumber.Text = "1"

        clockhour = 0
        clockmin = 0
        clocksec = 0

        WindowsMediaPlayer1.URL = ""
    End Sub

'[] TIMERS []

    Public Sub TimerClock_Timer()
        clockhour = Hour(Time)
        clockmin = Minute(Time)
        clocksec = Second(Time)
        LabelClock.Caption = Format(clockhour, "00") & ":" & Format(clockmin, "00") & ":" & Format(clocksec, "00")
    End Sub

    Public Sub TimerDice_Timer()
        If dicecounter <= 0 Then
            TimerDice.Enabled = False
            dicecounter = 6
            FormMainWindow.MousePointer = 99
            If dicestatus < 3 Then MenuDiceRoll.Enabled = True: CmdDiceRoll.Enabled = True
            Exit Sub
        End If
        If dicecounter > 0 Then dicecounter = dicecounter - 1
        Call DiceRollOnce
    End Sub

'[] DOUBLER []

    Public Sub MenuDoublerX2_Click()
        doublernumberprev = doublernumber
        doublernumber = doublernumber * 2
        Call DoublerRefresher
    End Sub
    Public Sub CmdDoublerX2_Click()
        Call MenuDoublerX2_Click
    End Sub
    Public Sub MenuDoublerX3_Click()
        doublernumberprev = doublernumber
        doublernumber = doublernumber * 3
        Call DoublerRefresher
    End Sub
    Public Sub CmdDoublerX3_Click()
        Call MenuDoublerX3_Click
    End Sub
    Public Sub MenuDoublerX5_Click()
        doublernumberprev = doublernumber
        doublernumber = doublernumber * 5
        Call DoublerRefresher
    End Sub
    Public Sub CmdDoublerX5_Click()
        Call MenuDoublerX5_Click
    End Sub
    Public Sub MenuDoublerX6_Click()
        doublernumberprev = doublernumber
        doublernumber = doublernumber * 6
        Call DoublerRefresher
    End Sub
    Public Sub CmdDoublerX6_Click()
        Call MenuDoublerX6_Click
    End Sub
    Public Sub MenuDoublerUndo_Click()
        doublernumber = doublernumberprev
        Call DoublerRefresher
    End Sub
    Public Sub CmdDoublerUndo_Click()
        Call MenuDoublerUndo_Click
    End Sub
    Public Sub MenuDoublerReset_Click()
        doublernumber = 1
        doublernumberprev = 1
        Call DoublerRefresher
    End Sub
    Public Sub CmdDoublerReset_Click()
        Call MenuDoublerReset_Click
    End Sub

    Public Sub DoublerRefresher()
        Select Case doublernumber
            Case 1
                TextboxDoublerNumber.ForeColor = &HC0C0C0
            Case Is > 100
                TextboxDoublerNumber.ForeColor = &HFF&
            Case Else
                TextboxDoublerNumber.ForeColor = &HFF8000
        End Select

        If doublernumber > 99999 Then
            doublernumber = 99999
            Select Case setlanguage
                Case "ENG"
                    MsgBox "The maximum number of the doubler is 99999.", vbExclamation + vbOKOnly, "CAUTION"
                Case "CHS"
                    MsgBox "倍数的上限为 99999。", vbExclamation + vbOKOnly, "注意"
                Case "JPN"
                    MsgBox "倍数の最大値は 99999 です。", vbExclamation + vbOKOnly, "注意"
            End Select
        End If

        'JUDGE WHETHER UNDO IS AVAILABLE...
        If doublernumber = doublernumberprev Then
            MenuDoublerUndo.Enabled = False
            CmdDoublerUndo.Enabled = False
        Else
            MenuDoublerUndo.Enabled = True
            CmdDoublerUndo.Enabled = True
        End If

        TextboxDoublerNumber.Text = doublernumber
    End Sub

'[] DICE []

    Public Sub MenuDiceRoll_Click()
        If dicestatus = 3 Then Exit Sub
        If soundswitch = True Then WindowsMediaPlayer1.URL = App.Path & "\CZJST_appdata\CZJST_audio\CZJSTaudio_DiceRoll.wav"

        TimerDice.Enabled = True
        dicecounter = 6
        FormMainWindow.MousePointer = 11
        MenuDiceRoll.Enabled = False
        CmdDiceRoll.Enabled = False
    End Sub
    Public Sub CmdDiceRoll_Click()
        Call MenuDiceRoll_Click
    End Sub
    Public Sub MenuDiceReset_Click()
        dicestatus = 0
        dicenumber1 = 0
        dicenumber2 = 0
        dicetemp = 0
        dicecounter = 0
        WindowsMediaPlayer1.URL = ""
        Call DiceRefresher
    End Sub
    Public Sub CmdDiceReset_Click()
        Call MenuDiceReset_Click
    End Sub

    Public Sub DiceRollOnce()
        Select Case dicestatus
            Case 0
                dicestatus = 1
            Case 1
                Call DiceNumberFormer
                dicenumber1 = dicetemp
                Call DiceRefresher  'PAY ATTENTION TO THE ORDER!
                If dicecounter = 0 Then
                        TextboxDiceNumber1.ForeColor = &H80FF&
                        dicestatus = 2
                End If
            Case 2
                Call DiceNumberFormer
                dicenumber2 = dicetemp
                If dicecounter = 0 Then
                    'ANTI-REPEAT!
                    If dicenumber2 = dicenumber1 Then
                        dicecounter = 1
                    Else
                        TextboxDiceNumber2.ForeColor = &H80FF&
                        dicestatus = 3
                    End If
                End If
                Call DiceRefresher  'PAY ATTENTION TO THE ORDER!
            Case Else
                'THIS IS AN ERROR!
                Select Case setlanguage
                    Case "ENG"
                        MsgBox "Sorry, an error has occurred and the program has partly stopped working. We would appreciate it if you can send a feedback to us so as to help solve the problem.", vbCritical + vbOKOnly, "ERROR"
                    Case "CHS"
                        MsgBox "很抱歉，程序发生异常。请向我们提供反馈。", vbCritical + vbOKOnly, "错误"
                    Case "JPN"
                        MsgBox "すみません、エラーが発生しました。著者に報告してください。", vbCritical + vbOKOnly, "エラー"
                End Select
        End Select
    End Sub
    Public Sub DiceNumberFormer()
        Randomize
        dicetemp = Int(13 * Rnd) + 1
    End Sub

    Public Sub DiceRefresher()
        TextboxDiceNumber1.Text = dicenumber1
        TextboxDiceNumber2.Text = dicenumber2

        Select Case dicestatus
            Case 0
                TextboxDiceNumber1.ForeColor = &HC0C0C0
                TextboxDiceNumber2.ForeColor = &HC0C0C0
                MenuDiceRoll.Enabled = True
                CmdDiceRoll.Enabled = True
                Select Case setlanguage
                    Case "ENG"
                        MenuDiceRoll.Caption = "START"
                    Case "CHS"
                        MenuDiceRoll.Caption = "开始"
                    Case "JPN"
                        MenuDiceRoll.Caption = "スタート"
                End Select
            Case 1
                TextboxDiceNumber1.ForeColor = &HC0C0C0
                TextboxDiceNumber2.ForeColor = &HC0C0C0
                Select Case setlanguage
                    Case "ENG"
                        MenuDiceRoll.Caption = "CONTINUE"
                    Case "CHS"
                        MenuDiceRoll.Caption = "继续"
                    Case "JPN"
                        MenuDiceRoll.Caption = "続く"
                End Select
            Case 2
                TextboxDiceNumber1.ForeColor = &H80FF&
                TextboxDiceNumber2.ForeColor = &HC0C0C0
                Select Case setlanguage
                    Case "ENG"
                        MenuDiceRoll.Caption = "CONTINUE"
                    Case "CHS"
                        MenuDiceRoll.Caption = "继续"
                    Case "JPN"
                        MenuDiceRoll.Caption = "続く"
                End Select
            Case 3
                TextboxDiceNumber1.ForeColor = &H80FF&
                TextboxDiceNumber2.ForeColor = &H80FF&
                MenuDiceRoll.Enabled = False
                CmdDiceRoll.Enabled = False
                Select Case setlanguage
                    Case "ENG"
                        MenuDiceRoll.Caption = "FINISHED"
                    Case "CHS"
                        MenuDiceRoll.Caption = "已完成"
                    Case "JPN"
                        MenuDiceRoll.Caption = "完了しました"
                End Select
        End Select

        CmdDiceRoll.Caption = MenuDiceRoll.Caption

        'USE POKER ALPHABET TO REPLACE NUMBER...
        If dicenumber1 = 0 Then TextboxDiceNumber1.Text = "?"
        If dicenumber1 = 1 Then TextboxDiceNumber1.Text = "A"
        If dicenumber1 = 11 Then TextboxDiceNumber1.Text = "J"
        If dicenumber1 = 12 Then TextboxDiceNumber1.Text = "Q"
        If dicenumber1 = 13 Then TextboxDiceNumber1.Text = "K"
        If dicenumber2 = 0 Then TextboxDiceNumber2.Text = "?"
        If dicenumber2 = 1 Then TextboxDiceNumber2.Text = "A"
        If dicenumber2 = 11 Then TextboxDiceNumber2.Text = "J"
        If dicenumber2 = 12 Then TextboxDiceNumber2.Text = "Q"
        If dicenumber2 = 13 Then TextboxDiceNumber2.Text = "K"

        'CHANGE TITLE TEXT...
        If TextboxDiceNumber2.Text = "?" Then
            Select Case setlanguage
                Case "ENG"
                    LabelDiceNumber1.Caption = "Dice"
                    LabelDiceNumber2.Caption = "-"
                Case "CHS"
                    LabelDiceNumber1.Caption = "癞子"
                    LabelDiceNumber2.Caption = "-"
                Case "JPN"
                    LabelDiceNumber1.Caption = "サイコロ"
                    LabelDiceNumber2.Caption = "-"
            End Select
        Else
            Select Case setlanguage
                Case "ENG"
                    LabelDiceNumber1.Caption = "DiceA"
                    LabelDiceNumber2.Caption = "DiceB"
                Case "CHS"
                    LabelDiceNumber1.Caption = "天癞子"
                    LabelDiceNumber2.Caption = "地癞子"
                Case "JPN"
                    LabelDiceNumber1.Caption = "サイコロＡ"
                    LabelDiceNumber2.Caption = "サイコロＢ"
            End Select
        End If
    End Sub

'[] MENU []

    'CMD General...
    Public Sub MenuEXIT_Click()
        End
    End Sub

    'CMD Language...
    Public Sub MenuLanguageENG_Click()
        Call ModuleLoadLanguage.LoadLanguageENG
    End Sub
    Public Sub MenuLanguageCHS_Click()
        Call ModuleLoadLanguage.LoadLanguageCHS
    End Sub
    Public Sub MenuLanguageJPN_Click()
        Call ModuleLoadLanguage.LoadLanguageJPN
    End Sub

    'CMD Preferences...
    Public Sub MenuSoundSwitch_Click()
        If soundswitch = True Then
            soundswitch = False
            Select Case setlanguage
                Case "ENG"
                    MenuSoundSwitch.Caption = "Soun&d OFF"
                Case "CHS"
                    MenuSoundSwitch.Caption = "声音 关 (&D)"
                Case "JPN"
                    MenuSoundSwitch.Caption = "音声 オフ (&D)"
            End Select
            WindowsMediaPlayer1.URL = ""
        Else
            soundswitch = True
            Select Case setlanguage
                Case "ENG"
                    MenuSoundSwitch.Caption = "Soun&d ON"
                Case "CHS"
                    MenuSoundSwitch.Caption = "声音 开 (&D)"
                Case "JPN"
                    MenuSoundSwitch.Caption = "音声 オン (&D)"
            End Select
            WindowsMediaPlayer1.URL = ""
        End If
    End Sub

'================================================================================

'================================================================================
