#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

CoordMode Screen

OrderBoxX = 130
OrderBoxY = 191


SaveMousePos() {
    WinActivate, Trio SCS - Acctivate
    MouseGetPos, OrderBoxX, OrderBoxY
}


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


MoveBack() {
    MouseMove, OrderBoxX, OrderBoxY
}


    ;creates a new note in acctivate
    ^!a::
    NewNote()
    return

    ;Creates a PO receiving note to both buyer and prompts for note to Anne
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
    FormatTime, month,, MM
    FormatTime, day,, dd
    Send %month%{right}%day%{right}
    Send ^s
    ifNotEqual noteAnne, y
        return

    ;Only sets note to Anne
    ^!q::
    NewNote()
    Send All Parts Recieved, PO Ready For Completion.
    Send {Tab}{Tab}{Tab}a{Down}{Tab}
    FormatTime, month,, MM
    FormatTime, day,, dd
    Send %month%{right}%day%{right}
    Send ^s
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
    FormatTime, month,, MM
    FormatTime, day,, dd
    Send %month%{right}%day%{right}
    Send ^s
    return


;Creates a ship note, prompts for # shipped
    ^!s::

    QtyPrompt := "Enter number of items shipped"
    InputBox, Qual, Enter Qual, %QtyPrompt%
    if ErrorLevel
        return

    NewNote()
    Send Picked and Pulled, Shipping %Qual% Pcs. Please Invoice and Send Confirmation Note.
    Send {Tab}{Tab}{Tab}a{Down}{Tab}
    FormatTime, month,, MM
    FormatTime, day,, dd
    Send %month%{right}%day%{right}
    Send ^s
    return

ParseXlCol(Xl, col, entryNum, arr) {
    /*Send {Ctrl Down}c{Ctrl Up}
    Sleep, 20
    StringReplace, clipboard, clipboard, `r`n, ,All
    value = %clipboard%
    return value
    */
    ;return Xl.Range("A1").Value
    while (Xl.Range("A" . A_Index).Value != "") {
        Xl.Range("A" . A_Index).Value := value
    }

}

;incomplete for shipping notes entry
/*
AutoShip() {
    WinGetTitle, currentExcel
    Xl := ComObjActive("Excel.Application") ;creates a handle to your currently active excel sheet

    QtyPrompt := "Enter number of Orders to process"
    InputBox, Qual, Enter Qual, %QtyPrompt%
    /*if ErrorLevel > 0
        Msgbox Input Error
        return
        */

    /*
    POArray := Object()
    POQty := Object()
    Loop, %Qual% {
        POArray.Insert(ParseClipboard())
        Send {Right}
        Sleep, 20
        POQty.Insert(ParseClipboard())
        Send {Left}{Down}
        Sleep, 20
    }

    WinActivate, Untitled - Notepad

    for index, element in POArray {
        Send "Element number " . %index% . " is " . %element%
        Send " " POQty[index]
        Send {`n}
    }

    /*
    WinActivate, Trio SCS - Acctivate
    WinWaitActive, Trio SCS - Acctivate

    MouseGetPos, OBoxX, OBoxY

    WinActivate, Invoice Notes - Excel
    WinWaitActive, Invoice Notes - Excel

    Send {Right}
    cellValue := ParseClipboard()
    Send {Left}

    POvalue := ParseClipboard()
    Send {Down}

    WinActivate, Trio SCS - Acctivate
    WinWaitActive, Trio SCS - Acctivate

    Click OBoxX, OBoxY
    Send {Home}
    Send +{End}
    Send %POvalue%{Tab}
}
    */

; Transfers from list of cells in Excel to ACCTIVATE
CellTransfer(loopCount) {

    loop, %loopCount% {
        Send {Ctrl Down}c{Ctrl Up}
        Sleep, 40
        Send {Down}
        Sleep, 20

        Send {Alt Down}{Tab}{Alt Up}
        Sleep, 80

        Send {Ctrl Down}v{Ctrl Up}
        Sleep, 30
        Send {Backspace}
        Sleep, 20
        Send {Down}
        Sleep, 20

        Send {Alt Down}{Tab}{Alt Up}
        Sleep, 80
    }
    
    return
}


    ^!c::
    QtyPrompt := "Enter number of items"
    InputBox, Qual, Enter Qual, %QtyPrompt%
    if ErrorLevel
        return

    CellTransfer(Qual)
    return

    ^!d::
    Xl := ComObjActive("Excel.Application")
    tester := ParseClipboard(Xl, "A",)
    MsgBox %tester%    
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