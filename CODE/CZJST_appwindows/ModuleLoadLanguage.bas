Attribute VB_Name = "ModuleLoadLanguage"
'================================================================================

'================================================================================

Public Sub LoadLanguageENG()
    FormMainWindow.setlanguage = "ENG"

    FormMainWindow.Caption = "Fight Landlord Card Game Assistant　v1.01　by Sam Toki"

    FormMainWindow.MenuLanguageENG.Checked = True
    FormMainWindow.MenuLanguageCHS.Checked = False
    FormMainWindow.MenuLanguageJPN.Checked = False

    FormMainWindow.MenuDoubler.Caption = "D&oubler"
    FormMainWindow.MenuDoublerUndo.Caption = "←　Undo"
    FormMainWindow.MenuDoublerReset.Caption = "＊　Reset"
    FormMainWindow.MenuDice.Caption = "D&ice"
    FormMainWindow.MenuDiceRoll.Caption = "N/A"
    FormMainWindow.MenuDiceReset.Caption = "＊　Reset"
    If FormMainWindow.soundswitch = True Then FormMainWindow.MenuSoundSwitch.Caption = "Soun&d ON" Else FormMainWindow.MenuSoundSwitch.Caption = "Soun&d OFF"
    FormMainWindow.MenuAbout.Caption = "&About"
    FormMainWindow.MenuEXIT.Caption = "E&XIT"

    FormMainWindow.FrameDoubler.Caption = "Doubler"
    FormMainWindow.FrameDoubler.Font = "Microsoft Sans Serif"
    FormMainWindow.FrameDice.Caption = "Dice"
    FormMainWindow.FrameDice.Font = "Microsoft Sans Serif"
    FormMainWindow.CmdDiceRoll.Font = "Microsoft Sans Serif"
    FormMainWindow.LabelDiceNumber1.Font = "Microsoft Sans Serif"
    FormMainWindow.LabelDiceNumber2.Font = "Microsoft Sans Serif"
    Call FormMainWindow.DoublerRefresher
    Call FormMainWindow.DiceRefresher
End Sub

'================================================================================

Public Sub LoadLanguageCHS()
    FormMainWindow.setlanguage = "CHS"

    FormMainWindow.Caption = "线下斗地主棋牌辅助工具　v1.01　Sam Toki 制作"

    FormMainWindow.MenuLanguageENG.Checked = False
    FormMainWindow.MenuLanguageCHS.Checked = True
    FormMainWindow.MenuLanguageJPN.Checked = False

    FormMainWindow.MenuDoubler.Caption = "倍数 (&O)"
    FormMainWindow.MenuDoublerUndo.Caption = "←　撤销"
    FormMainWindow.MenuDoublerReset.Caption = "＊　重置"
    FormMainWindow.MenuDice.Caption = "癞子 (&I)"
    FormMainWindow.MenuDiceRoll.Caption = "N/A"
    FormMainWindow.MenuDiceReset.Caption = "＊　重置"
    If FormMainWindow.soundswitch = True Then FormMainWindow.MenuSoundSwitch.Caption = "声音 开 (&D)" Else FormMainWindow.MenuSoundSwitch.Caption = "声音 关 (&D)"
    FormMainWindow.MenuAbout.Caption = "关于 (&A)"
    FormMainWindow.MenuEXIT.Caption = "退出 (&X)"

    FormMainWindow.FrameDoubler.Caption = "倍数"
    FormMainWindow.FrameDoubler.Font = "SimSun"
    FormMainWindow.FrameDice.Caption = "癞子"
    FormMainWindow.FrameDice.Font = "SimSun"
    FormMainWindow.CmdDiceRoll.Font = "SimHei"
    FormMainWindow.LabelDiceNumber1.Font = "SimHei"
    FormMainWindow.LabelDiceNumber2.Font = "SimHei"
    Call FormMainWindow.DoublerRefresher
    Call FormMainWindow.DiceRefresher
End Sub

'================================================================================

Public Sub LoadLanguageJPN()
    FormMainWindow.setlanguage = "JPN"

    FormMainWindow.Caption = "Fight Landlord カードゲームアシスタント　v1.01　by Sam Toki"

    FormMainWindow.MenuLanguageENG.Checked = False
    FormMainWindow.MenuLanguageCHS.Checked = False
    FormMainWindow.MenuLanguageJPN.Checked = True

    FormMainWindow.MenuDoubler.Caption = "番数 (&O)"
    FormMainWindow.MenuDoublerUndo.Caption = "←　取り消す"
    FormMainWindow.MenuDoublerReset.Caption = "＊　リセット"
    FormMainWindow.MenuDice.Caption = "サイコロ (&I)"
    FormMainWindow.MenuDiceRoll.Caption = "N/A"
    FormMainWindow.MenuDiceReset.Caption = "＊　リセット"
    If FormMainWindow.soundswitch = True Then FormMainWindow.MenuSoundSwitch.Caption = "音声 オン (&X)" Else FormMainWindow.MenuSoundSwitch.Caption = "音声 オフ (&D)"
    FormMainWindow.MenuAbout.Caption = "について (&A)"
    FormMainWindow.MenuEXIT.Caption = "终了 (&X)"

    FormMainWindow.FrameDoubler.Caption = "番数"
    FormMainWindow.FrameDoubler.Font = "MS UI Gothic"
    FormMainWindow.FrameDice.Caption = "サイコロ"
    FormMainWindow.FrameDice.Font = "MS UI Gothic"
    FormMainWindow.CmdDiceRoll.Font = "MS UI Gothic"
    FormMainWindow.LabelDiceNumber1.Font = "MS UI Gothic"
    FormMainWindow.LabelDiceNumber2.Font = "MS UI Gothic"
    Call FormMainWindow.DoublerRefresher
    Call FormMainWindow.DiceRefresher
End Sub

'================================================================================

'================================================================================
