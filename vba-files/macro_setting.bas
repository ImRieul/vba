Attribute VB_Name = "macro_setting"

public sub setting()

    ' main.checkCellUseColor
        Application.MacroOptions Macro:="main.checkCellUseColor", Description:="", _
        ShortcutKey:="C"

    ' main.mergeWithShortCut
        Application.MacroOptions Macro:="main.mergeWithShortCut", Description:="", _
        ShortcutKey:="Q"

    ' main.mergeGroup
        Application.MacroOptions Macro:="main.mergeGroup", Description:="", _
        ShortcutKey:="G"

    ' main.tableDivision
        Application.MacroOptions Macro:="main.tableDivision", Description:="", _
        ShortcutKey:="D"


end sub