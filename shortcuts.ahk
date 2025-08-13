#Requires AutoHotkey v2.0
#SingleInstance
Persistent ; kinda unnecessary, but it's good to have it (important for continuous activation)

TraySetIcon ".\rocket.ico", , 1

A_IconTip := "Shortcuts" ; Tooltip appears on hovering the tray icon.

;?  Tipp: powershell kann mit: Run "pwsh" gestartet werden

/*
;? Checking for Drive-letter
    copiedPath := A_Clipboard
    driveLetter := SubStr(copiedPath, 1, 2)
*/


; MENU
Tray := A_TrayMenu
Tray.Delete() ; Vordefinierte Menüpunkte löschen.
Tray.Add("Welcome to Shortcuts!", WelcomeUI)
Tray.Add() ; Trennlinie
Tray.Add("Check Keeping Alive", CKA)
Tray.Add("Toggle Startup", ToggleStartup)
Tray.Add()
Tray.Add("Quit", QuitApp)

WelcomeUI(*) {
    MyGui := Gui()
    MyGui.BackColor := "White"
    MyGui.Add("Picture", "x0 y0 h350 w450", A_WinDir "\Web\Wallpaper\Windows\img0.jpg")
    MyBtn := MyGui.Add("Button", "Default xp+20 yp+250", "Start the Deletion")
    MyBtn.OnEvent("Click", MoveBar)
    MyProgress := MyGui.Add("Progress", "w416")
    MyText := MyGui.Add("Text", "wp")  ; wp means "use width of previous".
    MyGui.Show()

    MoveBar(*)
    {
        Loop Files, A_WinDir "\*.*", "R"
        {
            if (A_Index > 100)
                break
            MyProgress.Value := A_Index
            MyText.Value := A_LoopFileName
            Sleep 50
        }
        MyText.Value := "Deletion complete."
    }
    MyGui.OnEvent("Close", CloseUI)
    CloseUI(*) {
        
        MsgBox("Welcome to Shortcuts, a small program for helping with everyday tasks.")
        MsgBox("All Shortcuts have all Keybinds: Press 'Alt' and a number")
        TrayTip "Shortcuts", "Thank you for using Shortcuts", "Icon Mute"
    }



}
CKA(*) {
    if TimerRunning
        TrayTip "Shortcuts", "Keep Alive is active", "Iconi Mute"
    else
        TrayTip "Shortcuts", "Keep Alive is not active", "Iconx Mute"
}
ToggleStartup(*){
    targetPath := A_Startup "\" StrReplace(A_ScriptName, ".ahk", ".lnk")

    if FileExist(targetPath) {
        ; Wenn schon im Autostart, dann entfernen
        FileDelete(targetPath)
        ; TrayTip "Shortcuts", "Autostart deaktiviert", "IconX Mute" 
        ShowPopup("Autostart deaktiviert")
    }
    else {
        ; Wenn noch nicht im Autostart, dann kopieren
        shell := ComObject("WScript.Shell")
        shortcut := shell.CreateShortcut(targetPath)
        shortcut.TargetPath := A_ScriptFullPath
        shortcut.WorkingDirectory := RegExReplace(A_ScriptFullPath, "\\[^\\]+$")
        shortcut.Save()
        ; TrayTip "Shortcuts", "Autostart aktiviert", "IconI Mute"
        ShowPopup("Autostart aktiviert")
    }

}
QuitApp(*){
    ExitApp()
}

ShowPopup(text) {
    Popup := Gui(, "Shortcuts")
    Popup.Opt("+AlwaysOnTop -Caption +ToolWindow")
    Popup.SetFont("s14 w500") ; Schriftgröße 10, dicke
    Popup.Add("Text", "cacacac", text) 
    Popup.BackColor := "5c035c"
    ; Popup.OnEvent("Click", Popup.Destroy())
    Popup.Show("AutoSize xCenter Y50")

    ; custom Window mit Abgerundete Ecken
    Popup.GetPos(&x, &y, &w, &h)
    radius := 35  ; Radius in Pixeln
    region := DllCall("CreateRoundRectRgn", "Int", 0, "Int", 0, "Int", w, "Int", h, "Int", radius, "Int", radius, "Ptr")
    DllCall("SetWindowRgn", "Ptr", Popup.Hwnd, "Ptr", region, "Int", true)
   
    ; Destroy Popup with a Click
    WM_LBUTTONDOWN := 0x0201
    handler := (wParam, lParam, msg, hwnd) => (hwnd = Popup.Hwnd ? (Popup.Destroy(), 0) : 0)
    OnMessage(WM_LBUTTONDOWN, handler)

    try { ;if the Popup wasn't clicked
        SetTimer(() => Popup.Destroy(), -3000)
    }
     
}

configFile := A_ScriptDir "\shortcutsConfig.ini"

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
            
            ; wird getoggelt, wenn: 0 und StartupFile da || 1 und StartupFile nicht da
            targetPath := A_Startup "\" StrReplace(A_ScriptName, ".ahk", ".lnk")
            if (FileExist(targetPath) && Startup == 0) || (!FileExist(targetPath) && Startup == 1) {
                ToggleStartup()
            }

            if (Startup == "") {
                ; ini existiert ist aber leer
                IniWrite false, configFile, "Startup", "AutoStart"
            }
        }
    }




!1:: { ;Open VsCode for this folder
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

!2:: { ;Open VsCode in Ordner
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

!3:: { ;Open IntelliJ in Ordner
    if WinActive("ahk_exe explorer.exe") {
        Send "^l"
        Sleep 200
        Send "^c"
        Sleep 100
        Run "C:\WINDOWS\system32\cmd.exe"
        if WinWait("C:\WINDOWS\system32\cmd.exe", , 3) {
            WinActivate

            Send "idea64.exe `"" A_Clipboard "`"{Enter}"
            WinWait("ahk_exe idea64.exe", , 5)
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

; Setting for Mouse Detection
LockTimerRunning := false
checkInterval := 100       ; Prüfintervall in Millisekunden
lastX := 0
lastY := 0

!9:: {
    global LockTimerRunning, lastX, lastY
    LockTimerRunning := !LockTimerRunning
    if LockTimerRunning {
        TrayTip "Shortcuts", "Maus-Überwachung AKTIV", "Iconi Mute"
        SetTimer () => TrayTip(), -1000
        MouseGetPos(&lastX, &lastY)
        SetTimer(CheckMouseMove, checkInterval)
    } else {
        TrayTip "Shortcuts", "Maus-Überwachung INAKTIV", "Iconx Mute"
        SetTimer(CheckMouseMove, 0)
        SetTimer () => TrayTip(), -1000
    }
    Sleep(800)
    ToolTip("")
}




TimerRunning := false

!0:: { ; Keep Alive

    global TimerRunning
    TimerRunning := !TimerRunning  ; Toggle Zustand

    if (TimerRunning) {
        TraySetIcon ".\favicon.ico", , 
        A_IconTip := ""
        SetTimer PressNumLock, 180000  ; 180.000 ms = 3 Minuten
        TrayTip "Shortcuts", "Keep Alive started", "Iconi Mute"
        ; ToolTip "Fn-Loop aktiviert (alle 5 Minuten)"
        SetTimer () => TrayTip(), -3000  ; TrayTip nach 3 Sekunde ausblenden
    } else {
        TraySetIcon ".\rocket.ico", , 1
        A_IconTip := "Shortcuts"
        SetTimer PressNumLock, 0  ; Timer stoppen
        ; ToolTip "Fn-Loop deaktiviert"
        TrayTip "Shortcuts", "Keep Alive stopped", "Iconx Mute"
        SetTimer () => TrayTip(), -1000
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
