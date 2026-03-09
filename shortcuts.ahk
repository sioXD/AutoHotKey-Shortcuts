#Requires AutoHotkey v2.0
#SingleInstance
Persistent ; kinda unnecessary, but it's good to have it (important for continuous activation)

TraySetIcon ".\images\rocket.ico", , 1

A_IconTip := "Shortcuts" ; Tooltip appears on hovering the tray icon.
configFile := A_ScriptDir "\shortcutsConfig.ini" ; Config File

; Read Hotkeys
global HK_CodeWorkspace := IniRead(configFile, "Hotkeys", "CodeWorkspace", "!1")
global HK_CodeFolder := IniRead(configFile, "Hotkeys", "CodeFolder", "!2")
global HK_NeoVimFolder := IniRead(configFile, "Hotkeys", "NeoVimFolder", "!3")
global HK_GitClone := IniRead(configFile, "Hotkeys", "GitClone", "!4")
global HK_WebSearch := IniRead(configFile, "Hotkeys", "WebSearch", "!7")
global HK_AutoLogin := IniRead(configFile, "Hotkeys", "AutoLogin", "!8")
global HK_MouseDetection := IniRead(configFile, "Hotkeys", "MouseDetection", "!9")
global HK_KeepAlive := IniRead(configFile, "Hotkeys", "KeepAlive", "!0")

RegisterHotkeys() {
    global
    try Hotkey HK_CodeWorkspace, Action_CodeWorkspace, "On"
    try Hotkey HK_CodeFolder, Action_CodeFolder, "On"
    try Hotkey HK_NeoVimFolder, Action_NeoVimFolder, "On"
    try Hotkey HK_GitClone, Action_GitClone, "On"
    try Hotkey HK_WebSearch, Action_WebSearch, "On"
    try Hotkey HK_AutoLogin, Action_AutoLogin, "On"
    try Hotkey HK_MouseDetection, Action_MouseDetection, "On"
    try Hotkey HK_KeepAlive, Action_KeepAlive, "On"
}
UnregisterHotkeys() {
    global
    try Hotkey HK_CodeWorkspace, "Off"
    try Hotkey HK_CodeFolder, "Off"
    try Hotkey HK_NeoVimFolder, "Off"
    try Hotkey HK_GitClone, "Off"
    try Hotkey HK_WebSearch, "Off"
    try Hotkey HK_AutoLogin, "Off"
    try Hotkey HK_MouseDetection, "Off"
    try Hotkey HK_KeepAlive, "Off"
}
RegisterHotkeys()

;?  Tipp: powershell kann mit: Run "pwsh" gestartet werden

/*
;? Checking for Drive-letter
    copiedPath := A_Clipboard
    driveLetter := SubStr(copiedPath, 1, 2)
*/


; MENU
Tray := A_TrayMenu
Tray.Delete() ; Vordefinierte Menüpunkte löschen.
Tray.Add("Settings", SettingsUI)
Tray.Add("Reset Hotkeys", ResetHotkeys)
Tray.Add() ; Trennlinie
Tray.Add("Check Keeping Alive", CKA)
Tray.Add("Toggle Startup", ToggleStartup)
Tray.Add("Toggle AutoServerStart", ToggleAutoServerStart)
Tray.Add("Change AutoServerStart Credentials", ChangeCredentials)
Tray.Add()
Tray.Add("Quit", QuitApp)

SettingsUI(*) {
    global
    UnregisterHotkeys() ; Disable hotkeys while editing
    MyGui := Gui(, "Settings")
    MyGui.OnEvent("Close", (*) => RegisterHotkeys()) ; Re-enable if closed without saving
    
    MyGui.BackColor := "White"
    MyGui.SetFont("s10", "Segoe UI")

    MyGui.Add("Text", "w400 x20 y20", "Shortcuts - Settings:")
    
    MyGui.Add("Text", "x20 y+15 w200", "Open VsCode Workspace:")
    MyGui.Add("Hotkey", "x+10 yp-3 w120 vHK_CodeWorkspace", HK_CodeWorkspace)
    
    MyGui.Add("Text", "x20 y+15 w200", "Open VsCode in Ordner:")
    MyGui.Add("Hotkey", "x+10 yp-3 w120 vHK_CodeFolder", HK_CodeFolder)
    
    MyGui.Add("Text", "x20 y+15 w200", "Open NeoVim in Ordner:")
    MyGui.Add("Hotkey", "x+10 yp-3 w120 vHK_NeoVimFolder", HK_NeoVimFolder)
    
    MyGui.Add("Text", "x20 y+15 w200", "git clone from Clipboard:")
    MyGui.Add("Hotkey", "x+10 yp-3 w120 vHK_GitClone", HK_GitClone)
    
    MyGui.Add("Text", "x20 y+15 w200", "Websearch selected Text:")
    MyGui.Add("Hotkey", "x+10 yp-3 w120 vHK_WebSearch", HK_WebSearch)
    
    MyGui.Add("Text", "x20 y+15 w200", "AutoLogin on Website:")
    MyGui.Add("Hotkey", "x+10 yp-3 w120 vHK_AutoLogin", HK_AutoLogin)
    
    MyGui.Add("Text", "x20 y+15 w200", "Mouse Detection Toggle:")
    MyGui.Add("Hotkey", "x+10 yp-3 w120 vHK_MouseDetection", HK_MouseDetection)
    
    MyGui.Add("Text", "x20 y+15 w200", "Keep Alive Toggle:")
    MyGui.Add("Hotkey", "x+10 yp-3 w120 vHK_KeepAlive", HK_KeepAlive)
    
    MyGui.SetFont("s9 cGray")
    MyGui.Add("Text", "x20 y+15 w200 Center", "<To Disable Press 'Space'>")
    MyGui.SetFont("s10 cDefault")

    MyGui.Add("Text", "x20 y+20 w400", "--------------------------------------------------------")

    ;* Autostart toggle
    targetPath := A_Startup "\" StrReplace(A_ScriptName, ".ahk", ".lnk")

    opt := "x20 y+10 w400"
    if FileExist(targetPath)
        opt .= " Checked"

    chkAutostart := MyGui.Add("CheckBox", opt " vChk_Autostart", " Enable Autostart on Boot")


    MyBtnSave := MyGui.Add("Button", "w100 x110 y+30 Default", "Save")
    MyBtnSave.OnEvent("Click", SaveShortcuts)

    MyBtnClose := MyGui.Add("Button", "w100 x+20 yp", "Close")
    MyBtnClose.OnEvent("Click", CloseSettings)

    MyGui.Show("w400 h500")

    CloseSettings(*) {
        RegisterHotkeys()
        MyGui.Destroy()
    }

    SaveShortcuts(*) {
        Saved := MyGui.Submit()

        ; Apply new
        HK_CodeWorkspace := Saved.HK_CodeWorkspace
        HK_CodeFolder := Saved.HK_CodeFolder
        HK_NeoVimFolder := Saved.HK_NeoVimFolder
        HK_GitClone := Saved.HK_GitClone
        HK_WebSearch := Saved.HK_WebSearch
        HK_AutoLogin := Saved.HK_AutoLogin
        HK_MouseDetection := Saved.HK_MouseDetection
        HK_KeepAlive := Saved.HK_KeepAlive

        ; Save to INI
        IniWrite(HK_CodeWorkspace, configFile, "Hotkeys", "CodeWorkspace")
        IniWrite(HK_CodeFolder, configFile, "Hotkeys", "CodeFolder")
        IniWrite(HK_NeoVimFolder, configFile, "Hotkeys", "NeoVimFolder")
        IniWrite(HK_GitClone, configFile, "Hotkeys", "GitClone")
        IniWrite(HK_WebSearch, configFile, "Hotkeys", "WebSearch")
        IniWrite(HK_AutoLogin, configFile, "Hotkeys", "AutoLogin")
        IniWrite(HK_MouseDetection, configFile, "Hotkeys", "MouseDetection")
        IniWrite(HK_KeepAlive, configFile, "Hotkeys", "KeepAlive")

        ; Handle Autostart toggle if changed
        targetPath := A_Startup "\" StrReplace(A_ScriptName, ".ahk", ".lnk")
        if (Saved.Chk_Autostart == 1 && !FileExist(targetPath)) || (Saved.Chk_Autostart == 0 && FileExist(targetPath)) {
            ToggleStartup()
        }

        RegisterHotkeys()
        ShowPopup("Settings saved!", "001d2b", "039590", 1500)
    }
}

ResetHotkeys(*) {
    global
    ; Turn off currently active hotkeys
    UnregisterHotkeys()

    ; Delete the Hotkeys config section
    try IniDelete configFile, "Hotkeys"

    ; Re-assign default hotkeys
    HK_CodeWorkspace := "!1"
    HK_CodeFolder := "!2"
    HK_NeoVimFolder := "!3"
    HK_GitClone := "!4"
    HK_WebSearch := "!7"
    HK_AutoLogin := "!8"
    HK_MouseDetection := "!9"
    HK_KeepAlive := "!0"

    ; Re-register the defaults
    RegisterHotkeys()
    
    ShowPopup("Hotkeys reset to default", "001d2b", "039590", 1500)
}

CKA(*) {
    if TimerRunning
        TrayTip "Shortcuts", "Keep Alive is active", "Iconi Mute"
    else
        TrayTip "Shortcuts", "Keep Alive is not active", "Iconx Mute"
}
ToggleAutoServerStart(*){
    if !FileExist(configFile) 
        || !(AutoStartServer := IniRead(configFile, "Startup", "AutoServerStart", ""))
        ||  (AutoStartServer := IniRead(configFile, "Startup", "AutoServerStart", "")) == false
    {
        IniWrite true, configFile, "Startup", "AutoServerStart"
    } else {
        IniWrite false, configFile, "Startup", "AutoServerStart"
    }
}

ChangeCredentials(*){
    url := InputBox("Browser URL eingeben:", "url").Value
    username := InputBox("Benutzername eingeben:", "Login").Value
    password := InputBox("Passwort eingeben:", "Login", "Password").Value ; Passwort-Eingabe maskiert

    ; Speichern in einer INI-Datei
    IniWrite url, configFile, "AutoServerStart", "url"
    IniWrite username, configFile, "AutoServerStart", "Username"
    IniWrite password, configFile, "AutoServerStart", "Password"

}



ToggleStartup(*){
    targetPath := A_Startup "\" StrReplace(A_ScriptName, ".ahk", ".lnk")

    if FileExist(targetPath) {
        ; Wenn schon im Autostart, dann entfernen
        FileDelete(targetPath)
        IniWrite false, configFile, "Startup", "AutoStart"
        ; TrayTip "Shortcuts", "Autostart deaktiviert", "IconX Mute" 
        ShowPopup("Autostart deaktiviert")
    } else {
        ; Wenn noch nicht im Autostart, dann kopieren
        shell := ComObject("WScript.Shell")
        shortcut := shell.CreateShortcut(targetPath)
        shortcut.TargetPath := A_ScriptFullPath
        shortcut.WorkingDirectory := RegExReplace(A_ScriptFullPath, "\\[^\\]+$")
        shortcut.Save()
        IniWrite true, configFile, "Startup", "AutoStart"
        ; TrayTip "Shortcuts", "Autostart aktiviert", "IconI Mute"
        ShowPopup("Autostart aktiviert")
    }

}
QuitApp(*){
    ExitApp()
}

ShowPopup(text, color:="FAE492", background:="2f4858", time:=3000) { 
    if WinExist("ShortcutsPopup ahk_class AutoHotkeyGUI") {
        ; bestehendes Popup zerstören
        WinKill("ShortcutsPopup ahk_class AutoHotkeyGUI")
    }

    Popup := Gui("-Caption", "ShortcutsPopup")
    Popup.Opt("+AlwaysOnTop -Caption +ToolWindow +E0x08000000")
    Popup.SetFont("s15 w600", "Candara Code") ; Schriftgröße dicke , Font
    Popup.Add("Text", "c" color, text) 
    Popup.BackColor := background

    ;! Popup.Show("NA AutoSize x0 y0")
    Popup.Show("NA AutoSize xCenter yCenter")
    Popup.GetPos(,, &w, &h)
    MonitorGetWorkArea(MonitorGetPrimary(), &L, &T, &R, &B)
    

    ; Move to bottom-right corner
    newX := R - w - 10  ; 10px margin from edge
    newY := B - h - 10
    ;! Popup.Move(newX, newY)


    ; custom Window mit Abgerundete Ecken
    Popup.GetPos(&x, &y, &w, &h)
    radius := 35  ; Radius in Pixeln
    region := DllCall("CreateRoundRectRgn", "Int", 0, "Int", 0, "Int", w, "Int", h, "Int", radius, "Int", radius, "Ptr")
    DllCall("SetWindowRgn", "Ptr", Popup.Hwnd, "Ptr", region, "Int", true)


   
    ; Destroy Popup with a Click
    WM_LBUTTONDOWN := 0x0201
    
    handler(wParam, lParam, msg, hwnd) {
        try {
            if hwnd = Popup.Hwnd {
                Popup.Destroy()
                OnMessage(WM_LBUTTONDOWN, handler, 0) ; Handler wieder entfernen
            }
        } catch {
            ; Fenster existiert nicht mehr, nichts tun
            OnMessage(WM_LBUTTONDOWN, handler, 0)
        }
        return 0
    }

    OnMessage(WM_LBUTTONDOWN, handler)

    try { ;if the Popup wasn't clicked
        SetTimer(() => Popup.Destroy(), -time)
    }
     
}

; On StartUp
; Prüfen, ob Konfigurationsdatei schon existiert
if !FileExist(configFile) {
    targetPath := A_Startup "\" StrReplace(A_ScriptName, ".ahk", ".lnk")

    shell := ComObject("WScript.Shell")
    shortcut := shell.CreateShortcut(targetPath)
    shortcut.TargetPath := A_ScriptFullPath
    shortcut.WorkingDirectory := RegExReplace(A_ScriptFullPath, "\\[^\\]+$")
    shortcut.Save()
    ; TrayTip "Shortcuts", "Autostart aktiviert", "IconI Mute"
    ShowPopup("Autostart aktiviert")
    IniWrite true, configFile, "Startup", "AutoStart"
    
} else {
    try{
        Startup := IniRead(configFile, "Startup", "AutoStart", "")
        AutoStartServer := IniRead(configFile, "Startup", "AutoStartServer", "")
        
        ; wird getoggelt, wenn: 0 und StartupFile da || 1 und StartupFile nicht da
        targetPath := A_Startup "\" StrReplace(A_ScriptName, ".ahk", ".lnk")
        if (FileExist(targetPath) && Startup == 0) || (!FileExist(targetPath) && Startup == 1) {
            ToggleStartup()
        }

        if (Startup == "") {
            ; ini existiert ist aber leer
            IniWrite false, configFile, "Startup", "AutoStart"
        }

        if (AutoStartServer) {
            Send "!8"
        }

    }
}



Action_CodeWorkspace(ThisHotkey) { ;Open VsCode for this folder
    Run "C:\WINDOWS\system32\cmd.exe"
    if WinWait("C:\WINDOWS\system32\cmd.exe", , 3) {
        WinActivate
        Sleep 200
        Send "code `"" A_ScriptDir "`"{Enter}"
        Sleep 500
        WinWait("ahk_exe Code.exe", , 5)
        ;! maybe wait for specific window: example from spy => [shortcuts.ahk - ahk - Visual Studio Code] -> "ahk" is the folder
        loop { ; force kill
            if WinExist("C:\WINDOWS\system32\cmd.exe") {
                WinKill "C:\WINDOWS\system32\cmd.exe"
                Sleep 100
            } else
                break
        }

    } else {
        MsgBox "couldn't run the Program"
    }
}

Action_CodeFolder(ThisHotkey) { ;Open VsCode in Ordner
    if WinActive("ahk_exe explorer.exe") {
        Send "^l"
        Sleep 200
        Send "^c"
        Sleep 100
        Run "C:\WINDOWS\system32\cmd.exe"
        if WinWait("C:\WINDOWS\system32\cmd.exe", , 3) {
            WinActivate

            Send "code `"" A_Clipboard "`" {Enter}"
            Sleep 500
            WinWait("ahk_exe Code.exe", , 5)
            loop { ; force kill
                if WinExist("C:\WINDOWS\system32\cmd.exe") {
                    WinKill "C:\WINDOWS\system32\cmd.exe"
                    Sleep 100
                } else
                    break
            }

        } else {
            MsgBox "couldn't run the Program"
        }
    }
}

Action_NeoVimFolder(ThisHotkey) { ;Open NeoVim in Ordner
    if WinActive("ahk_exe explorer.exe") {
        Send "^l"
        Sleep 200
        Send "^c"
        Sleep 100
        Run "C:\WINDOWS\system32\cmd.exe"
        if WinWait("C:\WINDOWS\system32\cmd.exe", , 3) {
            WinActivate

            Send "nvim `"" A_Clipboard "`"{Enter}"
            

        } else {
            MsgBox "couldn't run the Program"
        }
    }
}

Action_GitClone(ThisHotkey) { ; git clone from Clipboard —e
    if !WinActive("ahk_exe explorer.exe")
        return

    repo := A_Clipboard

    Send "^l"
    Sleep 200
    Send "^c"
    Sleep 100


    cmd := A_ComSpec " /c git clone `"" repo "`""
    exitCode := RunWait(cmd, A_Clipboard, "Hide") ; WorkingDir gesetzt

    if (exitCode = 0) {
        ShowPopup("git clone finished", "001d2b", "039590", 1500)
    } else {
        MsgBox("git clone failed. Exit code: " exitCode)
    }
}


; Sebsearch with selected Text
Action_WebSearch(ThisHotkey) {
   selectedText := GetSelectedText()
    if (selectedText)
        Run("https://www.google.com/search?q=" . selectedText)

}



; AutoLogin on Website
ServerTimerRunning := false
Action_AutoLogin(ThisHotkey) {
    global ServerTimerRunning
    ServerTimerRunning := !ServerTimerRunning


    ; get server
    if !FileExist(configFile) 
        || !(url := IniRead(configFile, "AutoServerStart", "url", "")) 
        || !(username := IniRead(configFile, "AutoServerStart", "Username", "")) 
        || !(password := IniRead(configFile, "AutoServerStart", "Password", "")) 
    {
        url := InputBox("Browser URL eingeben:", "url").Value
        username := InputBox("Benutzername eingeben:", "Login").Value
        password := InputBox("Passwort eingeben:", "Login", "Password").Value ; Passwort-Eingabe maskiert

        ; Speichern in einer INI-Datei
        IniWrite url, configFile, "AutoServerStart", "url"
        IniWrite username, configFile, "AutoServerStart", "Username"
        IniWrite password, configFile, "AutoServerStart", "Password"
    } else {
        ; Aus Datei lesen
        url := IniRead(configFile, "AutoServerStart", "url", "")
        username := IniRead(configFile, "AutoServerStart", "Username", "")
        password := IniRead(configFile, "AutoServerStart", "Password", "")
    }
    ; domain bekommen
    if RegExMatch(url, "^(?:https?:\/\/)?(?:[^@\/\n]+@)?(?:www\.)?([^:\/?\n]+)", &match) {
        server := match[1]
    } else {
        MsgBox "Keine Domain gefunden. \nShortcut wird beendet."
        ServerTimerRunning := false
    }


    if (ServerTimerRunning) {
        SetTimer(CheckServer(url, server, username, password), 3000)
        ; More Beautiful
        TraySetIcon ".\images\lens.ico", , 1
        A_IconTip := "Lens Search"
    } else {
        SetTimer(CheckServer(url,server, username, password), 0)
        TraySetIcon ".\images\rocket.ico", , 1
        A_IconTip := "Shortcuts" 
    }
    
    

}




; Setting for Mouse Detection
LockTimerRunning := false
checkInterval := 100       ; Prüfintervall in Millisekunden
lastX := 0
lastY := 0

Action_MouseDetection(ThisHotkey) {
    global LockTimerRunning, lastX, lastY
    LockTimerRunning := !LockTimerRunning
    if LockTimerRunning {
        ; TrayTip "Shortcuts", "Maus-Überwachung AKTIV", "Iconi Mute"
        ; SetTimer () => TrayTip(), -1000
        ShowPopup("Mouse-Detection ACTIVE", "001d2b", "6fae8a")
        MouseGetPos(&lastX, &lastY)
        SetTimer(CheckMouseMove, checkInterval)
    } else {
        ; TrayTip "Shortcuts", "Maus-Überwachung INAKTIV", "Iconx Mute"
        ; SetTimer () => TrayTip(), -1000
        SetTimer(CheckMouseMove, 0)
        ShowPopup("Mouse-Detection INACTIVE", "001d2b", "be5845")
    }
}




TimerRunning := false

Action_KeepAlive(ThisHotkey) { ; Keep Alive

    global TimerRunning
    TimerRunning := !TimerRunning  ; Toggle Zustand

    if (TimerRunning) {
        TraySetIcon ".\images\favicon.ico", , 
        A_IconTip := ""
        SetTimer PressNumLock, 180000  ; 180.000 ms = 3 Minuten
        ; TrayTip "Shortcuts", "Keep Alive started", "Iconi Mute"
        ShowPopup("Keep Alive started", "001d2b", "039590", 1000)
        ; SetTimer () => TrayTip(), -3000  ; TrayTip nach 3 Sekunde ausblenden
    } else {
        TraySetIcon ".\images\rocket.ico", , 1
        A_IconTip := "Shortcuts" 
        SetTimer PressNumLock, 0  ; Timer stoppen 
        ; TrayTip "Shortcuts", "Keep Alive stopped", "Iconx Mute"
        ; SetTimer () => TrayTip(), -1000
        ShowPopup("Keep Alive stopped", "001d2b", "be5845", 1000)
    }

}


GetSelectedText() {
    ClipSaved := ClipboardAll()
    A_Clipboard := ""
    SendEvent "^c"
    SelectedText := A_Clipboard
    A_Clipboard := ClipSaved
    return SelectedText
}

;Funktion for AutoServer Start
CheckServer(url, server, username, password) {
    global ServerTimerRunning
     exitCode := RunWait(A_ComSpec " /c ping -n 1 " server " >nul", , "Hide")

    if (exitCode = 0) {
        ; Server erreichbar
        Run url

        Sleep 3000 ; Warten, bis Seite geladen ist

        Send username   ; Benutzername eingeben
        Send A_Tab      ; Zum Passwortfeld springen
        Send password   ; Passwort eingeben
        Send "{Enter}"  ; Formular absenden
        
        Sleep 1000

        ; TODO for later
/*             #Include lib\WebScrapping.ahk ; for DOM
        scraper := WebScrapping()

        scraper := WebScrapping("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
        Sleep(1000)

        scraper.SetPageByURL(server)

        winTitle := WinGetTitle("A")
        MsgBox "Aktives Fenster: " winTitle

        Loop {
            el := scraper.GetElement("document.querySelector('.storeapp-list')")
            if (el) {
                MsgBox "Login erfolgreich, storeapp-list gefunden!"
                break
            }
            Sleep 500
        }
*/


        SetTimer(CheckServer(url,server, username, password), 0)
        
    } else {
        ; TrayTip "Lens Search", "keine Verbindung", "Icon! Mute"
        ToolTip "searching"
        SetTimer () => TrayTip(), -1000
    

        Sleep(2000)
    }
        
    


}

; Funktion: Mausbewegung prüfen
CheckMouseMove() {
    global lastX, lastY, LockTimerRunning
    MouseGetPos(&x, &y)
    if (x != lastX || y != lastY) {
        TrayTip "Shortcuts", "Bewegung erkannt", "Icon! Mute" 
        ; Sleep(500) ; 0.5 sek Pause, um ToolTip zu sehen
        
        ; deactivate the program
        LockTimerRunning := !LockTimerRunning
        SetTimer(CheckMouseMove, 0)

        DllCall("user32\LockWorkStation") ; call API for Locking windows
    }
}



PressNumLock() {
    Key := Random(1, 4)
    ; {NumLock} - {Ctrl} - {Shift} - {ScrollLock}
    switch Key {
        case 1:
            Key := "{NumLock}"
        case 2:
            Key := "{Ctrl}"
        case 3:
            Key := "{Shift}"
        case 4:
            Key := "{ScrollLock}"
        Default: ; if any error is error
            Key := "{NumLock}"
    }

    Sleep Random(10, 60000) ; Kurze Verzögerung (10ms - 1min)
    Send Key   ; 1. Drücken
    Sleep Random(10, 50) ; Kurze Verzögerung (10ms - 50ms)
    Send Key   ; 2. Drücken
    ; ToolTip Key
    ; SetTimer () => ToolTip(), -1000 ; Tooltip nach 1 Sekunde ausblenden
}
