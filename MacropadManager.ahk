#Requires AutoHotkey v1.1.31+  ; Requires recent AHK_L to use 'switch/case' functionality
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance, Force
SetTitleMatchMode, 2
CoordMode, Mouse, Screen
;=========================================
; Set tray icons
if (A_IsCompiled)
    Menu, Tray, Icon, %A_ScriptFullPath%, 1
; Set icon for Windows 10
If (SubStr(A_OSVersion,1,3) = "10.")
    Menu, Tray, Icon, C:\WINDOWS\system32\imageres.dll, 251	;Set custom Script icon
;=========================================
;Include dependencies from library
#Include %A_ScriptDir%\lib
#Include Ini.ahk
; Include personal files - ignored by git
#Include %A_ScriptDir%\personal

; Connect to INI configuration file using Ini.ahk class
IniFilePath:=A_ScriptDir "\MM_Settings.ini"
global Config:=Ini(IniFilePath)
global savedTabNumber:=0
InitiateHotkeys()
; CreateGUI() ; Optional - Open GUI on Startup
TrayTip(,"Macropad Manager running.")
Return


!F1::
^F19::
CreateGUI()
return

; Ctrl+Backspace deletes entire previous word rather than creating a [] character.
#If WinActive(A_ScriptName "ahk_class AutoHotkeyGUI")
^BackSpace::Send ^+{Left}{Backspace}
#If


CreateGui(){
    ; Initialise default GUI settings
    global
    Config:=Ini(IniFilePath)
    OnMessage(0x203,"WM_LButtonDBLCLK")
    ; OnMessage(0x200, "WM_MOUSEMOVE")  ; Used for debugging purposes to show control name when hovering.
    if WinExist(A_ScriptName " ahk_class AutoHotkeyGUI")
    {
        WinActivate
        return
    }
    oGui:={font:{ name:     "Segoe UI"
                , size:     10
                , color:    "Black"
                , styles:   "Norm" }
         , color:       "F5F5F5"
         , margin:      5
         , hDefault:    24
         , wKeyName:    18
         , wEditBox:    150
         , keyXGap:     ""
         , keyYGap:     "" }
    ; Refer to existing vars in object. Set these values separately
    oGui.keyXGap:=oGui.Margin*3, oGui.keyYGap:=Round((oGui.hDefault)/2)

    Gui, Main:New, +DPIScale +HwndgHwnd
    oGui.hwnd:=gHwnd
    Gui, Margin, % oGui.margin, % oGui.margin
    Gui, Font, % "s" oGui.font.size " c" oGui.font.color " " oGui.font.styles, % oGui.font.name
    Gui, Color, % oGui.color
    local tabHeadings:=""
    local garbageCleanup:=0
    ; Iterate over each keymapping to create tab names
    for index, mapName in Config.KeyMaps_Names
    {
        local displayName:=Config["KeyMaps_" mapName]["KeyMapDisplayName"]
        if !(displayName) ; If No display name is set, then ignore
        {
            ; Garbage collection
            MsgBox, 0x4, , % "Error: Key map '" mapName "' data was incomplete - no display name was set. Do you want to delete it?"
            IfMsgBox, % "Yes"
            {
                Config.KeyMaps_Names.Delete(index)
                Config.Delete("KeyMaps_" mapName)
                Sleep 100
                garbageCleanup++
            }
            IfMsgBox, % "No"
                MsgBox % "OK, but key map is kept hidden. Manually check config ini file to fix it."
        }
        else
            tabHeadings.= (displayName?displayName:"") "|"
    }
    ; Add an additional tab for creating new controls.
    tabHeadings.="+"
    Gui, Add, Tab3, % "vTab HwndtHwnd AltSubmit", % tabHeadings
    ; If a key mapping has been removed, then refresh the Config ini file
    if (garbageCleanup)
        Config:=Ini(IniFilePath)
    ; Iterate over each keymapping to create GUI Controls
    local baseKeyMap:=Config.KeyMaps_Names.1  ; By default, baseKeyMap = "Default"
    for index, mapName in Config.KeyMaps_Names
    {
        Gui, Tab, % index
        ; Create initial GUI elements on the selected tab – Edit and DDL with explanatory text.
        Gui, Add, Text, % "x+10 y+10 h" oGui.hDefault " Section", % "Window Name:"
        ; Status will be saved to: "Gui_KeyMaps_Default_WinCheckType", etc
        Gui, Add, Edit, % "x+m yp, h" oGui.hDefault " w200" (mapName=baseKeyMap?" ReadOnly":"") " vGui_KeyMaps_" mapName "_WinName", % Config["KeyMaps_" mapName]["WinName"]
        Gui, Add, Text, % "x+15 yp h" oGui.hDefault, % "Window Check Type:"
        if (mapName=baseKeyMap)
        {
            Gui, Add, Edit, % "x+m yp h" oGui.hDefault " w114" " ReadOnly" " vGui_KeyMaps_" mapName "_WinCheckType", % "N/A"
        }
        Else
        {
            local DDL:="IfWinActive|IfWinExist|"
            DDL:=StrReplace(DDL, (Config["KeyMaps_" mapName]["WinCheckType"]), (Config["KeyMaps_" mapName]["WinCheckType"] "|"))
            Gui, Add, DropDownList, % "x+m yp h" oGui.hDefault " w110" " vGui_KeyMaps_" mapName "_WinCheckType" " r3", % DDL
        }
        ; Default size and position of the group box. Will progressively grow as new elements added
        local xGroupBox:=oGui.Margin*2.5, yGroupBox:=oGui.hDefault*2+oGui.Margin
        ; Default size and position of the edit boxes and titles
        local keyXPos:=xGroupBox+oGui.Margin, keyXGap:=oGui.keyXGap, keyYGap:=oGui.keyYGap
        for each, layer in Config.Keys_Layers   ; Layer A and Layer B
        {
            local layerIndex:=each, local layerName:=layer
            local keyCount:=0
            local hGroupBox:=0, local wGroupBox:=keyXPos+oGui.Margin
            ; xGroupBox:=oGui.Margin*3, yGroupBox+=oGui.hDefault+hGroupBox, hGroupBox:=15
            ; First create the group box - will be resized later.
            Gui, Add, GroupBox, % "x" xGroupBox (layerIndex=1?" ys+" oGui.hDefault+oGui.Margin:" y+" oGui.hDefault) " Section vGroupBox_Tab" index "Layer" layerIndex, % "Layer " layerName
            local prevControl:="GroupBox"
            if (Config.Encoders.EnableEncoders)
            {
                if (Config.Encoders.EncoderClick)
                    local oEncoderVars:=[{name:"CCW",symbol:"⟲"},{name:"Click",symbol:"⌾"},{name:"CW",symbol:"⟳"}]
                Else
                    local oEncoderVars:=[{name:"CCW",symbol:"⟲"},{name:"CW",symbol:"⟳"}]
                Loop % Config.Encoders.NumEncoders
                {
                    local wGroupBox_Enc:=oGui.Margin*2.5
                    encNum:=A_Index
                    for key, val in oEncoderVars
                    {
                        ; Make the symbol size larger
                        Gui, Font, % "s" oGui.font.size*1.3
                        Gui, Add, Text, % "h" oGui.hDefault " w" oGui.wKeyName (key=1?" x" keyXPos:" x+" keyXGap) (key=1?(prevControl="GroupBox"?" ys+" oGui.hDefault : " y+" keyYGap):" yp") " Center" (key=1?" Section":""), % val.symbol
                        ; Status will be saved to: "Gui_KeyMaps_Default_LayerA_Enc1_CCW", etc
                        Gui, Font, % "s" oGui.font.size
                        Gui, Add, Edit, % "h" oGui.hDefault " w" oGui.wEditBox " x+m yp" " vGui_KeyMaps_" mapName "_Layer" layerName "_Enc" encNum "_" val.name, % Config["KeyMaps_" mapName]["Layer" layerName "_Enc" encNum "_" val.name]
                        wGroupBox_Enc+=oGui.Margin+oGui.wKeyName+oGui.Margin+oGui.wEditBox+oGui.Margin
                        prevControl:="Edit"
                    }
                    hGroupBox+=oGui.hDefault+keyYGap
                }
                ; Add Horizontal divider
                Gui, Add, Text, % "x" xGroupBox " y+" keyYGap/2 "  w" wGroupBox_Enc+1 " 0x10 Section"
                prevControl:="Divider"
            }
            Loop % Config.Gui.GridY {
                local gridRow:=A_Index
                local wGroupBox:=oGui.Margin*2.5
                Loop % Config.Gui.GridX {
                    local gridCol:=A_Index
                    keyCount++
                    ; m(gridRow, gridCol, keycount, "h80 w80" (gridCol=1?" xm":" x+m") (gridCol=1?" y+m":" yp"))
                    ; Set keyYPos (more legible than a bunch of nested ternary statements)                
                    if (prevControl="GroupBox")
                        keyYPos:="s+" oGui.hDefault
                    else if (prevControl="Divider")
                        keyYPos:="s+" keyYGap/2
                    else
                        if gridCol=1
                            keyYPos:="+" keyYGap
                        else
                            keyYPos:="p"
                    Gui, Add, Text, % "h" oGui.hDefault " w" oGui.wKeyName (gridCol=1?" x" keyXPos:" x+" keyXGap)  " y" keyYPos " Center", % keycount
                    ; Status will be saved to: "Gui_KeyMaps_Default_LayerA_1", etc
                    Gui, Add, Edit, % "h" oGui.hDefault " w" oGui.wEditBox " x+m yp" " vGui_KeyMaps_" mapName "_Layer" layerName "_" keyCount, % Config["KeyMaps_" mapName]["Layer" layerName "_" keyCount]
                    wGroupBox+=oGui.Margin+oGui.wKeyName+oGui.Margin+oGui.wEditBox+oGui.Margin
                    prevControl:="Edit"
                }
                hGroupBox+=oGui.hDefault+keyYGap
            }
            ; wGroupBox+= (oGui.Margin * (Config.Gui.GridX) / 3)
            hGroupBox+=oGui.Margin*4
            ; wGroupBox+=oGui.Margin
            ; Re-size Groupbox - Note, if the encoder controls are wider than the regular controls, it will use the larger width.
            GuiControl, Move, % "GroupBox_Tab" index "Layer" layerIndex, % "w" (wGroupBox>wGroupBox_Enc?wGroupBox:wGroupBox_Enc) " h" hGroupBox
        }
        local hGuiTotal:=yGroupBox + hGroupBox + oGui.Margin
        local wButton:=80
        local yGap:=12
        local xGap:=(wGroupBox-(wButton*2))/3
        local xPosButton1:=xGap+oGui.Margin
        local xPosButton2:=xPosButton1+wButton+xGap
        Gui, Add, Button, % "y+" yGap " x" xPosButton1 " w" wButton " h" oGui.hDefault*1.2 " gMainButtonSave" " Section", % "&Save"
        Gui, Add, Button, % "yp x" xPosButton2 " w" wButton " h" oGui.hDefault*1.2 " gMainButtonHide", % "&Hide"
    }
    ; 'Create new keymap' (+) tab (within braces so section can be collapsed in VS Code)
    {
        Gui, Tab, % Config.KeyMaps_Names.MaxIndex()+1
        Gui, Add, Text, % "x15 y+10 h" oGui.hDefault, % "Create new key mapping:"
        Gui, Add, Text, % "xp y+12 h" oGui.hDefault " w" wGroupBox-20, % "Key map display name:"
        Gui, Add, Edit, % "xp y+m, h" oGui.hDefault " w" wGroupBox-20 " vNew_KeyMap_Name", % ""
        Gui, Add, Text, % "xp y+12 h" oGui.hDefault " w" wGroupBox-20, % "Name of window during which key map should apply (refer to AHK WinTitle documentation):"
        Gui, Add, Edit, % "xp y+m, h" oGui.hDefault  " w" wGroupBox-20 " vNew_KeyMap_WinName", % ""
        Gui, Add, Text, % "xp y+12 h" oGui.hDefault, % "Window Check Type (i.e. if window is active, or simply exists):"
        Gui, Add, DropDownList, % "xp y+m h" oGui.hDefault " w120" " vNew_KeyMap_WinCheckType" " r3", % "IfWinActive||IfWinExist"
        Gui, Add, Button, % "ys x" xPosButton1 " w" wButton " h" oGui.hDefault*1.2 " gMainButtonSave", % "&Save"
        Gui, Add, Button, % "yp x" xPosButton2 " w" wButton " h" oGui.hDefault*1.2 " gMainButtonHide", % "&Hide"
    }
    hGuiTotal+=yGap+oGui.hDefault+oGui.Margin*1.5
    if (savedTabNumber)
    {
        GuiControl, Choose, Tab, % savedTabNumber
        savedTabNumber:=0
    }
    ; Gui, Show, % "h" hGuiTotal
    Gui, Show
    return


    MainButtonHide:
    MainGuiEscape:
    Gui, Destroy
    Return

    MainButtonSave:
    Gui, Main:Default
    Gui, Submit, NoHide
    Config:=Ini(IniFilePath)
    baseKeyMap:=""
    baseKeyMap:=Config.KeyMaps_Names.1  ; By default, baseKeyMap = "Default"
    ; Write to Config ini file
    ; Pause automatic syncing
    ; Config.Sync(False)
    for mapindex, mapName in Config.KeyMaps_Names
    {
        If !(mapName=baseKeyMap)
        {
            GuiControlGet, WinCheckType, , % Gui_KeyMaps_%mapName%_WinCheckType
            GuiControlGet, WinName, , % Gui_KeyMaps_%mapName%_WinName
            ; m(WinCheckType,WinName)
            Config["KeyMaps_" mapName]["WinCheckType"]:=WinCheckType
            Config["KeyMaps_" mapName]["WinName"]:=WinName
        }
        for layerindex, layerName in Config.Keys_Layers
        {
            if (Config.Encoders.EnableEncoders)
            {
                if (Config.Encoders.EncoderClick)
                    local aEncoderVars:=["CCW","Click","CW"]
                Else
                    local aEncoderVars:=["CCW","CW"]
                loop % Config.Encoders.NumEncoders
                {
                    encoderNum:=A_Index
                    for key, val in aEncoderVars
                    {
                        keyValue:=""
                        keyValue:=Gui_KeyMaps_%mapName%_Layer%layerName%_Enc%encoderNum%_%val%
                        Config["KeyMaps_" mapName]["Layer" layerName "_Enc" encoderNum "_" val]:=keyValue
                    }
                }
            }
            for keyIndex, keyName in Config["Keys_Layer" layerName]
            {
                keyValue:=""
                ; Retrieve from Gui_KeyMaps_Default_LayerA_1, etc
                ; For some reason, GuiControlGet doesn't work as intended...
                ; GuiControlGet, keyValue, , % Gui_KeyMaps_%mapName%_Layer%layerName%_%keyIndex%
                keyValue:=Gui_KeyMaps_%mapName%_Layer%layerName%_%keyIndex%
                ; d("keyValue: " keyValue, "ghwnd: " ghwnd, "mapname: " mapname, "layername: " layername, "keyIndex: " keyIndex, "keyname: " keyname, Gui_KeyMaps_%mapName%_Layer%layerName%_%keyIndex%)
                Config["KeyMaps_" mapName]["Layer" layerName "_" keyIndex]:=keyValue
            }
        }
    }
    ; Make sure changes to object get written to ini file
    ; Config.Sync(True)
    ; Config.Persist()
    ; Check if new keymap has been created:
    local newKM_DisplayName:="", newKM_WinName:="", newKM_WinCheckType:=""
    GuiControlGet, newKM_DisplayName, , New_KeyMap_Name
    GuiControlGet, newKM_WinName, , New_KeyMap_WinName
    GuiControlGet, newKM_WinCheckType, , New_KeyMap_WinCheckType
    If (newKM_DisplayName && !newKM_WinName) || (!newKM_DisplayName && newKM_WinName)
    {
        MsgBox % "Error: Error creating new key map. Both the key map display name and the name of the window must be provided."
    }
    Else if (newKM_DisplayName && newKM_WinName)
    {
        ; First, append a blank new line to ini file
        FileAppend, `n, % IniFilePath
        newKM_Index:=Config.KeyMaps_Names.Count()
        newKM_Index++
        ; 3rd program has name KeyMap_Program2, etc
        newKM_ProgramNo:=newKM_Index-1
        ; Check if it already exist
        if (Config["KeyMaps_Program" newKM_ProgramNo])
        {
            newKM_ProgramNo:=0
            ; If so, find an unused number within first 20 digits
            Loop 20
            {
                if IsObject(Config["KeyMaps_Program" A_Index])
                    Continue
                else
                    newKM_ProgramNo:=A_Index
            }
        }
        if !(newKM_ProgramNo)
        {
            MsgBox % "Error creating new program name in keymap settings."
        }
        newKM_Name:="Program" newKM_ProgramNo
        ; Pause automatic syncing
        ; Config.Sync(False)
        ; Create a new section in the config ini file.
        Config.KeyMaps_Names[newKM_Index]:=newKM_Name
        Config["KeyMaps_" newKM_Name]:={KeyMapDisplayName:newKM_DisplayName, WinCheckType:newKM_WinCheckType, WinName:newKM_WinName}
        for layerIndex, layerName in Config.Keys_Layers
        {
            for keyIndex, keyName in Config["Keys_Layer" layerName]
                Config["KeyMaps_" newKM_Name]["Layer" layerName "_" keyIndex]:=""
        }
        ; Make sure changes to object get written to ini file
        ; Config.Sync(True)
        ; Config.Persist()
        savedTabNumber:=Config.KeyMaps_Names.Count()
        Gui, Destroy
        CreateGui()
    }
    InitiateHotkeys()
    TrayTip()
    return
}

InitiateHotkeys(){
    ; Create a hotkey for each key in the Config ini file
    ; First, get details of each KeyMap Name (e.g. Default, Program1, etc)
    local baseKeyMap:=Config.KeyMaps_Names.1  ; By default, baseKeyMap = "Default"
    for mapIndex, mapName in Config.KeyMaps_Names
    {
        ; If the KeyMap is "Default" (i.e. the default), then no need to check IfWinActive details, etc.
        If !(mapName=baseKeyMap)
        {
            WinName:=Config["KeyMaps_" mapName]["WinName"]
            WinCheckType:=Config["KeyMaps_" mapName]["WinCheckType"]
        }
        ; Then, loop over each layer of keys (e.g. LayerA, LayerB)
        for layerIndex, layerName in Config.Keys_Layers
        {
            ; Get details of each key in each layer (e.g. LayerA_1 = F13, etc)
            for keyIndex, keyName in Config["Keys_Layer" layerName]
            {
                ; Then find the action associated with that key for the relevant key mapping
                ; i.e. Contained in Config.KeyMaps_Program1.LayerA_1 etc
                action:=Config["KeyMaps_" mapName]["Layer" layerName "_" keyIndex]
                if !(action)
                    continue
                keyMapDisplayName:=Config["KeyMaps_" mapName]["KeyMapDisplayName"]
                ; Create a function object "fn" passing the parameters to the "ExecuteAction" function
                fn:=Func("ExecuteAction").Bind(mapName,keyMapDisplayName,winName,keyIndex,keyName,action)
                ; Set window context settings - IfWinActive, etc
                If (mapName=baseKeyMap)  ; Not needed for base keymapping, i.e. "Default"
                    Hotkey, If
                Else
                {
                    if (WinCheckType = "IfWinActive")
                        Hotkey, IfWinActive, % WinName
                    Else if (WinCheckType = "IfWinExist")
                        Hotkey, IfWinExist, % WinName
                    Else
                        Hotkey, If
                }
                ; Bind the hotkey to the function object
                Hotkey, % keyName, % fn
            }
        }
    }
    ; Notify(0).AddWindow("Hotkeys Launched.")
}

ExecuteAction(mapName,keyMapDisplayName,winName,index,keyName,action){
    ; m("mapName: " mapName,"keyMapDisplayName: " keyMapDisplayName, "winName: " winName, "index: " index, "keyName: " keyName, "action: " action)
    ; n("mapName: " mapName "`n" "keyMapDisplayName: " keyMapDisplayName "`n" "winName: " winName "`n" "index: " index "`n" "keyName: " keyName "`n" "action: " action "`n")
    ; For reference
    ; mapName: Program1
    ; keyMapDisplayName: VS Code
    ; winName: Visual Studio Code ahk_exe Code.exe
    ; index: 1
    ; keyName: F13
    ; action: {F5}
    ;==================
    ; For commands where the entire action string will exactly match, use this switch:
    Switch action
    {
        case "reload":
        {
            Reload
            return
        }
    }
    ; Otherwise, use this one, which matches the first word, then parse the rest of the string as 'parameters'
    Switch StrSplit(action," ")[1]
    {
        ; Mouse clicks
        case "!click":
        {
            ; Set coord mode as follows:
            ; !click 500 200 c
            switch StrSplit(action," ")[4]
            {
                case "client", "c":
                    CoordMode, Mouse, Client
                case "window", "w":
                    CoordMode, Mouse, Window
                case "relative", "r":
                    CoordMode, Mouse, Relative
                default:
                    CoordMode, Mouse, Screen
            }
            MouseClick, Left, % StrSplit(action," ")[2], % StrSplit(action," ")[3]
            CoordMode, Mouse, Screen
        }
        ; Hold a key down while the hotkey is pressed.
        case "!hold":
        {
            HoldKey(StrSplit(action, " ", , 2)[2])
        }
        ; Search highlighted text
        case "!google":
            SearchText(searchQuery:="",searchEngine:="Google")
        case "!DDG":
            SearchText(searchQuery:="",searchEngine:="DDG")
        ; Call a function
        case "!func", "!fn":
        {
            functionString:=StrSplit(action, " ", , 2)[2]
            functionName:=StrSplit(functionString,",")[1]
            if IsFunc(functionName)
                DynFunc(functionString)
            Else
            {
                MsgBox % "Error: Specified function doesn't exist:`n" functionName
                return
            }
        }
        ; Call a hotkey from another script
        case "!hotkey", "!hk":
        {
            ; Wait for modifiers to be released, so as to not inadvertently trigger other hotkeys.
            WaitModifiersReleased()
            CallHotkey(StrSplit(action, " ", , 2)[2])
        }
        case "!script":
        {
            ; Note: This has to be finalised.
            actionString:=StrSplit(action, " ", , 2)[2]
            SplitPath, action, fileName, fileDir, fileExtension, fileNameNoExt, fileDrive
            m(fileName, fileDir, fileExtension, fileNameNoExt, fileDrive)
        }
        default:
        {
            WaitModifiersReleased()
            Send % action
        }
    }
    ; MsgBox % "FOR DEBUGGING`n-------------------------`n" "mapName: " mapName "`n" "keyMapDisplayName: " keyMapDisplayName "`n" "winName: " winName "`n" "index: " index "`n" "keyName: " keyName "`n" "action: " action)
    return
}

WM_LButtonDBLCLK() {
    ; Used to rename a tab name when double clicking on the control.
    ; Then launch the RenameKeymapGui() function.
    global tHwnd, ghwnd, vTab, Config
    MouseGetPos,,,,hCtrl, 2
    If (hCtrl = tHwnd) {
        Gui, Main:Default
        Gui, +OwnDialogs
        GuiControlGet, tabNo,,tab
        KM_name:=Config.KeyMaps_Names[tabNo]
        KM_displayName:=Config["KeyMaps_" KM_name]["KeyMapDisplayName"]
        ; m(tabNo,Config.KeyMaps_Names.Count())
        if (tabNo > Config.KeyMaps_Names.Count())
            return
        RenameKeymapGui(tabNo,KM_name,KM_displayName)
        ; MsgBox, % "You double-clicked tab " KM_displayName ".`nCurrent tab no: " Trim(tabNo) "."
    }
}

/*
; Used for debugging purposes. Hover over control to show a tooltime of the control name.
WM_MOUSEMOVE() {
    ; See here: https://www.autohotkey.com/docs/commands/Gui.htm#ExToolTip
    static CurrControl, PrevControl, _TT  ; _TT is kept blank for use by the ToolTip command below.
    CurrControl := A_GuiControl
    If (CurrControl <> PrevControl)
    {
        ToolTip  ; Turn off any previous tooltip.
        SetTimer, DisplayToolTip, 1000
        PrevControl := CurrControl
    }
    return

    DisplayToolTip:
    SetTimer, DisplayToolTip, Off
    ToolTip % CurrControl  ; The leading percent sign tell it to use an expression.
    SetTimer, RemoveToolTip, 3000
    return

    RemoveToolTip:
    SetTimer, RemoveToolTip, Off
    ToolTip
    return
}
*/

RenameKeymapGui(KM_number:="",KM_name:="",KM_displayName:=""){
    global
    if (!KM_number && !KM_Name)
    {
        MsgBox % "Error renaming keymapping. Must specify either keymap name or number."
        return
    }
    if !(KM_number)
    {
        for key, val in Config.KeyMaps_Names {
            if val = KM_name
            {
                KM_number:=val
                break
            }
        }
    }
    if !(KM_name)
        KM_name:=Config.KeyMaps_Names[KM_number]
    if !(KM_displayName)
        KM_displayName:=Config["KeyMaps_" KM_name]["KeyMapDisplayName"]
    if ! KM_number || !KM_Name || !KM_displayName {
        Msgbox % "Error getting details for keymap to rename."
        return
    }
    Gui, Rename:New
    local wText:=190, wEdit:=200, wTotal:=wText+wEdit+oGui.Margin
    Gui, Margin, % oGui.margin, % oGui.margin
    Gui, Font, % "s" oGui.font.size " c" oGui.font.color " " oGui.font.styles, % oGui.font.name
    Gui, Color, % oGui.color
    Gui, Add, Text, % "xm y+m w" wText " h" oGui.hDefault, % "Current Keymap Number|Name:"
    Gui, Add, Edit, % "x+m yp w" wEdit " h" oGui.hDefault " ReadOnly vcurrentKM_Name", % KM_number " | " KM_name
    Gui, Add, Text, % "xm y+m w" wText " h" oGui.hDefault, % "Current Keymap Display Name:"
    Gui, Add, Edit, % "x+m yp w" wEdit " h" oGui.hDefault " ReadOnly vcurrentKM_DisplayName", % KM_displayName
    Gui, Add, Text, % "xm y+m w" wText " h" oGui.hDefault, % "New Keymap Display Name:"
    Gui, Add, Edit, % "x+m yp w" wEdit " h" oGui.hDefault " vnewKM_DisplayName", % ""
    local wButton:=80
    local yGap:=10
    local xGap:=(wTotal-(wButton*2))/3
    local xPosButton1:=xGap+oGui.Margin
    local xPosButton2:=xPosButton1+wButton+xGap
    Gui, Add, Button, % "y+" yGap " x" xPosButton1 " w" wButton " h" oGui.hDefault " gButtonRenameSave", % "&Rename"
    Gui, Add, Button, % "yp x" xPosButton2 " w" wButton " h" oGui.hDefault " gButtonRenameCancel", % "&Cancel"
    GuiControl, Focus, % "newKM_DisplayName"
    Gui, Show, , % "Rename Keymap - " A_ScriptName
    Return

    ButtonRenameSave:
    Gui, Rename:Default
    GuiControlGet, tempKM_Name, , currentKM_Name
    currentKM_number:=StrSplit(tempKM_Name,"|")[1], currentKM_Name:=StrSplit(tempKM_Name," | ")[2]
    GuiControlGet, currentKM_DisplayName, , currentKM_DisplayName
    GuiControlGet, newKM_DisplayName, , newKM_DisplayName
    ; m(currentKM_Name,currentKM_DisplayName,newKM_DisplayName)
    ; m(KM_displayName,km_name,KM_number)
    Config["KeyMaps_" currentKM_Name]["KeyMapDisplayName"]:=newKM_DisplayName
    if (currentKM_number)
        savedTabNumber:=currentKM_number
    ; m(currentKM_Name,currentKM_number,savedTabNumber)
    Gui, Rename:Destroy
    Gui, Main:Destroy
    CreateGui()
    return

    ButtonRenameCancel:
    RenameGuiEscape:
    Gui, Destroy
    Return
}



;======================================
;       FUNCTIONS USED IN CODE
;======================================

DynFunc(funcAndParam,delim:=","){	;Usage: DynFunc("<function name>, <parameter 1>, ..., <parameter 9>")
    ; Used to dynamically call functions by passing a string to this function
    ; Credit: 0x00 from here: https://www.autohotkey.com/boards/viewtopic.php?t=62499
	p := StrSplit(funcAndParam,delim),_F:=p[1],pl := p.length()-1
	Loop 9
		_%A_Index%:=p[A_Index+1]
	IfGreater, pl, 9,MsgBox, 0x40010, %A_ScriptName%, % A_ThisFunc	": Too Many Paramters!`n`n" _F "(" StrReplace(funcAndParam,_F ",") " )"
	If !IsFunc(_F) || (pl > 9)
		Return
	
	( pl = 0 ? (r:=%_F%())
	: pl = 1 ? (r:=%_F%(_1))
	: pl = 2 ? (r:=%_F%(_1,_2))
	: pl = 3 ? (r:=%_F%(_1,_2,_3))
	: pl = 4 ? (r:=%_F%(_1,_2,_3,_4))
	: pl = 5 ? (r:=%_F%(_1,_2,_3,_4,_5))
	: pl = 6 ? (r:=%_F%(_1,_2,_3,_4,_5,_6))
	: pl = 7 ? (r:=%_F%(_1,_2,_3,_4,_5,_6,_7))
	: pl = 8 ? (r:=%_F%(_1,_2,_3,_4,_5,_6,_7,_8))
	: pl = 9 ? (r:=%_F%(_1,_2,_3,_4,_5,_6,_7,_8,_9))
	: )
	Return r
}

ParseHotkey(string,startSymbol:="[",endSymbol:="]"){
    ; Not actually used. Can probably delete.
    newString:=RegExReplace(string, "(LButton|RButton|MButton|XButton1|XButton2|WheelDown|WheelUp|WheelLeft|WheelRight|CapsLock|Space|Tab|Enter|Return|Escape|Esc|Backspace|BS|ScrollLock|Delete|Del|Insert|Ins|Home|End|PgUp|PgDn|Up|Down|Left|Right|NumLock|Numpad0|Numpad1|Numpad2|Numpad3|Numpad4|Numpad5|Numpad6|Numpad7|Numpad8|Numpad9|NumpadDot|NumpadDiv|NumpadMult|NumpadAdd|NumpadSub|NumpadEnter|NumpadIns|NumpadEnd|NumpadDown|NumpadPgDn|NumpadLeft|NumpadClear|NumpadRight|NumpadHome|NumpadUp|NumpadPgUp|NumpadDel|F1|F2|F3|F4|F5|F6|F7|F8|F9|F10|F11|F12|F13|F14|F15|F16|F17|F18|F19|F20|F21|F22|F23|F24|LWin|RWin|Control|Ctrl|Alt|Shift|LControl|LCtrl|RControl|RCtrl|LShift|RShift|LAlt|RAlt|Browser_Back|Browser_Forward|Browser_Refresh|Browser_Stop|Browser_Search|Browser_Favorites|Browser_Home|Volume_Mute|Volume_Down|Volume_Up|Media_Next|Media_Prev|Media_Stop|Media_Play_Pause|Launch_Mail|Launch_Media|Launch_App1|Launch_App2|AppsKey|PrintScreen|CtrlBreak|Pause|Break|Help|Sleep|sc\d{1,3}|vk\d{1,2})\b" , startSymbol "$1" endSymbol)
    return newString
}

CallHotkey(HotkeyToSend){
    ; Function to press a hotkey which can be intercepted by another script.
    backupSendLevel:=A_SendLevel
    SendLevel, 1
    Send, % HotkeyToSend
    Sleep 50
    SendLevel, % backupSendLevel
    return
}

WaitModifiersReleased(showMessage:=false){
    ; Release all modifier keys to avoid errors.
    ; Useful because my default keymapping includes keys like !+F13, etc.
    ; This function can probably be tidied up. Currently can go on for a long time if checking all modifiers...
	hotkeyString:=A_ThisHotkey
	errors:=0
    ; Remove left/right arrows. e.g. <^ = LCtrl.
	StrReplace(hotkeyString,"<")
	StrReplace(hotkeyString,">")
	if InStr(hotkeyString,"^"){
		KeyWait, Ctrl, 2
		errors+=ErrorLevel
	}
	if InStr(hotkeyString,"!"){
		KeyWait, Alt, 2
		errors+=ErrorLevel
	}
	if InStr(hotkeyString,"+"){
		KeyWait, Shift, 2
		errors+=ErrorLevel
	}
	if InStr(hotkeyString,"#"){
		KeyWait, LWin, 2
		errors+=ErrorLevel
		KeyWait, RWin, 2
		errors+=ErrorLevel
	}
	if (showMessage)
		MsgBox % "Number of KeyWait errors: " error
	return errors
}

TrayTip(title:="Macropad Manager",text:="Hotkeys successfully updated",duration:=2500,options:=0x30){
    ; Default options include a large icon and no notification sound.
    TrayTip, % title, % text, , 0x30
    if (duration)
    {
        Sleep % duration
        HideTrayTip()
    }
    return
}

HideTrayTip() {
    TrayTip  ; Attempt to hide tray tip the normal way.
    if SubStr(A_OSVersion,1,3) = "10." {
        Menu Tray, NoIcon
        Sleep 200  ; It may be necessary to adjust this sleep.
        Menu Tray, Icon
    }
    return
}

UriEncode(Uri){
	VarSetCapacity(Var, StrPut(Uri, "UTF-8"), 0)
	StrPut(Uri, &Var, "UTF-8")
	f := A_FormatInteger
	SetFormat, IntegerFast, H
	While Code := NumGet(Var, A_Index - 1, "UChar")
		If (Code >= 0x30 && Code <= 0x39 ; 0-9
         || Code >= 0x41 && Code <= 0x5A ; A-Z
         || Code >= 0x61 && Code <= 0x7A) ; a-z
			Res .= Chr(Code)
	Else
		Res .= "%" . SubStr(Code + 0x100, -1)
	SetFormat, IntegerFast, %f%
	Return, Res
}


HoldKey(keyName){
	Send % "{" keyName " Down}"
	KeyWait, % A_ThisHotkey
	Send % "{" keyName " Up}"
	return
}



;======================================
;   FUNCTIONS TO ALLOCATE TO MACROS
;======================================

SearchText(searchQuery:="",searchEngine:="Google",searchType:="",prepend:="",append:="",lucky:=False,browser:="Firefox",openURL:=True){
    if !(searchQuery)
    {
        searchQuery:=Clip()
        if !(searchQuery)
        {
            InputBox, searchQuery, Search query, , , 200, 100, , , , , %searchQuery%
            if ErrorLevel
                return
        }
    }
    if !(searchQuery)
        return
    ; If highlighted text is a URL, go straight there.
    ; Complex Regex pattern
    FoundPos:=RegExMatch(searchQuery, "^(http(s)?:\/\/.)?(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)$")
    if (FoundPos && openURL)
    {
        Run % searchQuery
        return
    }
    ; Otherwise, create the search URL and go there
    searchQuery:=UriEncode(Trim(searchQuery))
    fullSearchQuery:=(prepend?prepend "+":"") searchQuery (append?"+" append:"")
    switch searchEngine
    {
        case "Google", "G":
        {
            switch searchType
            {
                case "maps", "map", "mp":
                    URL:="http://www.google.com/maps/search/" fullSearchQuery
                case "images", "image", "i":
                    URL:="https://www.google.com/search?tbm=isch&q=" fullSearchQuery
                default:
                    URL:="https://www.google.com/search?q=" fullSearchQuery (lucky?"&btnI=1":"")
            }
        }
        case "DuckDuckGo","DDG":
        {
            switch searchType
            {
                case "maps", "mp":
                    URL:="https://duckduckgo.com/?iaxm=maps&q=" fullSearchQuery
                case "images", "i":
                    URL:="https://duckduckgo.com/?ia=images&iax=images&q=" fullSearchQuery
                default:
                    URL:="https://duckduckgo.com/?q=" (lucky?"!ducky+":"") fullSearchQuery
            }
        }
        default:
        {
            switch searchType
            {
                case "maps", "mp":
                    URL:="https://duckduckgo.com/?iaxm=maps&q=" fullSearchQuery
                case "images", "i":
                    URL:="https://duckduckgo.com/?ia=images&iax=images&q=" fullSearchQuery
                default:
                    URL:="https://duckduckgo.com/?q=" (lucky?"!ducky+":"") fullSearchQuery
            }
        }
    }
    if !(URL)
    {
        MsgBox % "Error creating URL to search"
        return
    }
    if (browser="Firefox")
        Run % "firefox.exe " URL
    else if (browser="Chrome")
        Run % "chrome.exe " URL
    else
        Run % "firefox.exe " URL
}



Reddit_OpenComments(){
	Send +c
	Sleep 500
	Send ^+{Tab}
}

Start_TogglTrack_Timer(){
    processName:="TogglDesktop.exe"
    winName:="Toggl Track"
    EnvGet, LocalAppData, LocalAppData
    processPath:=LocalAppData "\TogglDesktop\" processName
    If !(WinExist("ahk_exe " processName))
    {
        Run, % processPath, , , pid
        WinWait, % "ahk_exe " processName, , 3
        if (ErrorLevel)
        {
            MsgBox % "Error running " processName ".`n`nTerminating script."
            return
        }
        Sleep 500
        WinClose, % winName " ahk_exe " processName
        hwnd:=WinExist("ahk_exe " processName)
    }
    WinExist("ahk_exe " processName, , winName)
    WinActivate
    Send !{F1}
    return
}

;======================================
;   INCLUDE FUNCTIONS/LIBRARIES
;======================================

#Include, <WordLib>