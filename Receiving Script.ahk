#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

CoordMode Screen

OrderBoxX = 130
OrderBoxY = 191

Global setCompFlag = 0
Global customMonth = 0
Global customDay = 0

;Function: saves the current mouse position; not used
SaveMousePos() {
    WinActivate, Trio SCS - Acctivate
    MouseGetPos, OrderBoxX, OrderBoxY
}

;Function: moves the mouse back to position; not used
MoveBack() {
    MouseMove, OrderBoxX, OrderBoxY
}


;Function: Creates a new note in acctivate via imagesearch
NewNote()
{
    WinActivate, Trio SCS - Acctivate
    WinWaitActive, Trio SCS - Acctivate

    MouseGetPos, OrderBoxX, OrderBoxY

    ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, *100 P:\Warehouse\Jason backup\Migration\AutoIT\newbutt.bmp
    if (ErrorLevel=0)
    {
        ;MsgBox Image Found
        ;Sleep, 500
        Click %FoundX%, %FoundY%
    } else if (ErrorLevel=1) {
        MsgBox Image Not Found
        return
    } else {
        MsgBox Unknown Error
        Sleep, 500
        return
    }

    MouseMove, OrderBoxX, OrderBoxY

}

;Sends Pick and Pull Note to Anne
PickAndPull(Qual, POType) {
    NewNote()
    Send Picked and Pulled, Shipping %Qual% Pcs. Please Invoice and Send Confirmation Note.
    if (POType = "S")
        Send {Tab}{Tab}{Tab}a{Down}{Tab}
    else if (POType = "R")
        Send {Tab}{Tab}{Tab}j{Tab}
    else {
        MsgBox PickandPull Incorrect Input Error
        return
    }
    SetTimeCompSave()
    return
}

SetTimeCompSave() {
    if (customMonth != 0) {
        Send %customMonth%{right}%customDay%{right}
    } else {
        FormatTime, month,, MM
        FormatTime, day,, dd
        Send %month%{right}%day%{right}
    }
    if setCompFlag 
        Send {Tab}{Tab}{Space}
    Send ^s
}

; Transfers from list of cells in Excel to ACCTIVATE
CellTransfer(loopCount) {

    loop, %loopCount% {
        Send {Ctrl Down}c{Ctrl Up}
        Sleep, 50
        Send {Down}
        Sleep, 40

        Send {Alt Down}{Tab}{Alt Up}
        Sleep, 80

        Send {Ctrl Down}v{Ctrl Up}
        Sleep, 50
        Send {Backspace}
        Sleep, 40
        Send {Down}
        Sleep, 40

        Send {Alt Down}{Tab}{Alt Up}
        Sleep, 80
    }
    
    return
}

;creates a new note in acctivate
    ^!a::
    NewNote()
    return

;Creates a PO receiving note to both buyer and prompts for note to Anne
;Do not break order, no return statement to flow into ^!q command
    ^!w::
    qtyPrompt := "Enter a whole number. Do not use this to enter exceptions."
    InputBox, qty, Enter QTY received and accepted, %qtyPrompt%
    if ErrorLevel
        return

    BuyerPrompt := "Aniket = a, Dwi = d, Sam = s"
    InputBox, buyer, Enter Buyer, %BuyerPrompt%
    if ErrorLevel
        return

    InputBox, noteAnne, Complete PO?, y or n
    if ErrorLevel
        return

    NewNote()
    Send %qty% Received, %qty% Accepted.
    Send {Tab}{Tab}{Tab}%buyer%{Tab}
    SetTimeCompSave()
    ifNotEqual noteAnne, y
        return
    sleep, 300

;Only sets PO note to Anne
    ^!q::
    NewNote()
    Send All Parts Recieved, PO Ready For Completion.
    Send {Tab}{Tab}{Tab}a{Down}{Tab}
    SetTimeCompSave()
    return    


;Creates a producted arrived note, prompts for buyer
    ^!e::

    BuyerPrompt := "Aniket = a, Dwi = d, Sam = s"
    InputBox, buyer, Enter Buyer, %BuyerPrompt%
    if ErrorLevel
        return

    NewNote()
    Send Product Arrived, Not Yet Received.
    Send {Tab}{Tab}{Tab}%buyer%{Tab}
    SetTimeCompSave()
    return

;Creates a ship note, prompts for # shipped
    ^!s::

    QtyPrompt := "Enter number of items shipped"
    InputBox, Qual, Enter Qual, %QtyPrompt%
    if ErrorLevel
        return

    TypePrompt := "Sales order = S, TRO = R"
    InputBox, oType, Enter oType, %TypePrompt%
    if ErrorLevel
        return

    PickAndPull(Qual, oType)
    return

;Copy-Paste Shortcut, brute force
    ^!c::
    QtyPrompt := "Enter number of items"
    InputBox, Qual, Enter Qual, %QtyPrompt%
    if ErrorLevel
        return

    CellTransfer(Qual)
    return



;Work-in-Progress Starts Here

;takes in a com excel and a column (eg. "A")
;returns an array with the non-empty values of the column
ParseXlCol(Xl, col) {
    /*Send {Ctrl Down}c{Ctrl Up}
    Sleep, 20
    StringReplace, clipboard, clipboard, `r`n, ,All
    value = %clipboard%
    return value
    */
    ;return Xl.Range("A1").Value
    arr := object()

    while (Xl.Range(col . A_Index).Value != "") {
        ;MsgBox, % col . A_Index
        arr[A_Index] := Xl.Range(col . A_Index).Value
        ;MsgBox, % arr[A_Index]
    }
    return arr
}

;incomplete for shipping notes entry

AutoShip() {
    WinGetTitle, currentExcel
    Xl := ComObjActive("Excel.Application") ;creates a handle to your currently active excel sheet

    /*
    QtyPrompt := "Enter number of Orders to process"
    InputBox, Qual, Enter Qual, %QtyPrompt%
    if ErrorLevel
        return
    */

    WinActivate, Trio SCS - Acctivate

    MouseGetPos, OBoxX, OBoxY

    Xl := ComObjActive("Excel.Application")
    arrPO := ParseXlCol(Xl, "A")
    arrVar := ParseXlCol(Xl, "B")

    WinActivate, Trio SCS - Acctivate

    Loop, % arrPO.MaxIndex() {
        Click OBoxX, OBoxY
        Send {Home}
        Send +{End}
        Send, % arrPO[A_Index]
        Send {Tab}
        Sleep, 1500

        Ptype := SubStr(arrPO[A_Index],2,1)

        PickAndPull(floor(arrVar[A_Index]), Ptype)
        Sleep, 50
    }
}

    ^!d::
    AutoShip()
    return

    ^!1::
    if setCompFlag
        setCompFlag = 0
    else setCompFlag = 1
    MsgBox Completion Flag set at %setCompFlag%
    return

    ^!2::
    MonthPrompt := "Enter the month for entry (01-12), 0 for today"
    InputBox, customMonth, Enter customMonth, %MonthPrompt%
    if ErrorLevel
        return

    DayPrompt := "Enter the day for entry (01-31), 0 for today"
    InputBox, customDay, Enter customDay, %DayPrompt%
    if ErrorLevel
        return

    MsgBox Entry Date Set at %customMonth% / %customDay%
    return

;Creates a Tracking Note
TrackingNote() {
    QtyPrompt := "Enter number of items"
    InputBox, Qual, Enter Qual, %QtyPrompt%
    if ErrorLevel
        return

    TypePrompt := "Sales order = S, TRO = R"
    InputBox, oType, Enter oType, %TypePrompt%
    if ErrorLevel
        return

    Send {Alt Down}d{Alt up}{Right}
    NewNote()
    Send %Qual% Pcs Shipped VIA UPS Freight `n
    Send OCY 830588371
    if (oType = "S")
        Send {Tab}{Tab}{Tab}a{Down}{Tab}
    else if (oType = "R")
        Send {Tab}{Tab}{Tab}j{Tab}
    else {
        MsgBox Tracking Incorrect Input Error
        return
    }
    SetTimeCompSave()
    return
}

    ^!t::
    TrackingNote()
    return

/*
    ImageSearch, FoX, FoY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, *100 P:\Warehouse\Jason backup\Migration\AutoIT\sales_order_button.png
    if(ErrorLevel=0)
        {
        Click %FoX%, %FoY%
        Sleep, 500
        ImageSearch, FovX, FovY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, *70 P:\Warehouse\Jason backup\Migration\AutoIT\order_num.png
        if (ErrorLevel=0) {
            MsgBox Searching...
            Sleep, 500
            Click %OrderBoxX%, %FovY%
            OrderBoxX=%FovX%
            OrderBoxY=%FovY%
            Send %OrderBoxX% %OrderBoxY%
        } else if (ErrorLevel=1) {
            MsgBox OrderBox not found!
        } else {
            MsgBox Unknown Error!
        }
    }
*/