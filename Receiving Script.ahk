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
Global printFlag = 1

Global Quantity = 0
Global PONum = "TSO017"
Global Receiver = "tsli"
Global PartNo = "00Z000"
Global ShipCompany = "UPS Freight"
Global Tracking = "000000"
Global oType = ""

Global spec = 1
Global Tested = 0

;Function: saves the current mouse position; not used
SaveMousePos() {
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

ExportDoc() {
    WinActivate, Trio SCS - Acctivate
    WinWaitActive, Trio SCS - Acctivate

    MouseGetPos, OrderBoxX, OrderBoxY

    ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, *100 P:\Warehouse\Jason backup\Migration\AutoIT\exportbutton2.bmp
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

PrintDoc() {
    WinActivate, Trio SCS - Acctivate
    WinWaitActive, Trio SCS - Acctivate

    MouseGetPos, OrderBoxX, OrderBoxY

    ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, *100 P:\Warehouse\Jason backup\Migration\AutoIT\printbutton.bmp
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

CloseReportWin() {
    WinActivate, Trio SCS - Acctivate
    WinWaitActive, Trio SCS - Acctivate

    MouseGetPos, OrderBoxX, OrderBoxY

    ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, *100 P:\Warehouse\Jason backup\Migration\AutoIT\closereportwin2.bmp
    if (ErrorLevel=0)
    {
        FoundX+=60
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
    ;MouseMove, %FoundX%, %FoundY%
    ;MsgBox Ran
}

UploadButton() {
    MouseGetPos, OrderBoxX, OrderBoxY

    ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, *100 P:\Warehouse\Jason backup\Migration\AutoIT\smartvault.bmp
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
    if (Tested) {
        Send Picked and Pulled, IBM tested. Shipping %Qual% Pcs. Please Invoice and Send Confirmation Note.
    }
    else {
        Send Picked and Pulled, Shipping %Qual% Pcs. Please Invoice and Send Confirmation Note.
    }
    if (POType = "S")
        Send {Tab}{Tab}{Tab}a{Down}{Tab}
    else if (POType = "R")
        Send {Tab}{Tab}{Tab}j{Down}{Tab}
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

; Transfers from list of cells in Excel to ACCTIVATE; spec is 0 or 1
CellTransfer(loopCount, spec) {

    if spec
        loop, %loopCount% {
            Send {Ctrl Down}c{Ctrl Up}
            Sleep, 40
            Send {Right}
            Sleep, 40

            Send {Alt Down}{Tab}{Alt Up}
            Sleep, 100

            Send {Ctrl Down}v{Ctrl Up}
            Sleep, 50
            Send {Backspace}
            Sleep, 40
            Send {Right}
            Sleep, 40
            Send {Right}
            Sleep, 40

            Send {Alt Down}{Tab}{Alt Up}
            Sleep, 100

            Send {Ctrl Down}c{Ctrl Up}
            Sleep, 40
            Send {Down}
            Sleep, 40
            Send {Left}
            Sleep, 40

            Send {Alt Down}{Tab}{Alt Up}
            Sleep, 100

            Send {Ctrl Down}v{Ctrl Up}
            Sleep, 50
            Send {Backspace}
            Sleep, 40
            Send {Right}{Left}{Left}{Left}{Down}
            Sleep, 40

            Send {Alt Down}{Tab}{Alt Up}
            Sleep, 100
        }
    else 
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

    specPrompt := "Spec field? 1 for yes"
    InputBox, spec, Enter spec, %SpecPrompt%, , , , , , , ,%spec%
    if ErrorLevel
        return

    CellTransfer(Qual, spec)
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

AutoCopy() {
    WinGetTitle, currentExcel
    Xl := ComObjActive("Excel.Application") ;creates a handle to your currently active excel sheet

    /*
    QtyPrompt := "Enter number of Orders to process"
    InputBox, Qual, Enter Qual, %QtyPrompt%
    if ErrorLevel
        return
    */

    WinActivate, Trio SCS - Acctivate

    ;MouseGetPos, OBoxX, OBoxY

    Xl := ComObjActive("Excel.Application")
    arrPO := ParseXlCol(Xl, "A")
    arrVar := ParseXlCol(Xl, "B")

    WinActivate, Trio SCS - Acctivate

    Loop, % arrPO.MaxIndex() {
        ;Click OBoxX, OBoxY
        ;Send {Home}
        ;Send +{End}
        Send, % arrPO[A_Index]
        Sleep, 40
        Send {Right}{Right}
        Sleep, 40
        Send, % arrVar[A_Index]
        Sleep, 40
        Send {Right}{Left}{Left}{Left}{Down}
        Sleep, 40
    }
}

    ^!d::
    AutoShip()
    return

    ^!u::
    AutoCopy()
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

    ^!3::
    if printFlag
        printFlag = 0
    else printFlag = 1
    MsgBox Print Flag set at %printFlag%
    return

    ^!4::
    if Tested
        Tested = 0
    else Tested = 1
    MsgBox Tested flag set at %Tested%
    return

;Creates a Tracking Note
TrackingNote() {
    QtyPrompt := "Enter number of items"
    InputBox, Qual, Enter Qual, %QtyPrompt%
    if ErrorLevel
        return

    TypePrompt := "Sales order = S, TRO = R"
    InputBox, oType, Enter oType, %TypePrompt%, , , , , , , ,%oType%
    if ErrorLevel
        return

    ShipInfoPrompt := "Change Shipping Info? Enter y if yes."
    InputBox, shipPrompt, Enter ShipPrompt, %ShipInfoPrompt%
    if ErrorLevel
        return

    if (shipPrompt = "y") {
        CompanyPrompt := "Enter Shipping Method:"
        InputBox, ShipCompany, Enter ShipCompany, %CompanyPrompt%, , , , , , , ,%ShipCompany%
        if ErrorLevel
            return

        TrackingPrompt := "Enter Tracking Number:"
            InputBox, Tracking, Enter Tracking, %TrackingPrompt%, , , , , , , ,%Tracking%
            if ErrorLevel
        return
    }

    Send {Alt Down}d{Alt up}{Right}
    NewNote()
    Send %Qual% Pcs Shipped VIA %ShipCompany% `n
    Send %Tracking%
    if (oType = "S")
        Send {Tab}{Tab}{Tab}a{Down}{Tab}
    else if (oType = "R")
        Send {Tab}{Tab}{Tab}j{Down}{Tab}
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

SerialEnter() {
    QtyPrompt := "Enter number of serials"
    InputBox, Qual, Enter Qual, %QtyPrompt%
    if ErrorLevel
        return

    Loop, %Qual% {
        Click
        Sleep, 50
        Send {Enter}
        Sleep, 30
    }
}

    ^!k::
    SerialEnter()
    return

RMANote() {
    QtyPrompt := "Enter number of parts rejected"
    InputBox, Qual, Enter Qual, %QtyPrompt%
    if ErrorLevel
        return

    NewNote()
    Send %Qual% Pcs Rejected, Please RMA with Vendor
    Send {Tab}{Tab}{Tab}j{Down}{Tab}
    SetTimeCompSave()
}

    ^!f::
    RMANote()
    return

PrepOutbound() {
    POPrompt := "Enter the PO to work on:"
    InputBox, POnum, Enter POnum, %POPrompt%, , , , , , , ,%POnum%
    if ErrorLevel
        return

    QtyPrompt := "Enter number of items to pick:"
    InputBox, Qual, Enter Qual, %QtyPrompt%
    if ErrorLevel
        return

    SamePrompt := "Enter n if the different Part and Receiver:"
    InputBox, Same, Enter Same, %SamePrompt%
    if ErrorLevel
        return

    if (Same = "n") {
        PartPrompt := "Enter the part to be received. Enter 'Multi' if multiple parts:"
        InputBox, PartNo, Enter PartNo, %PartPrompt%, , , , , , , ,%PartNo%
        if ErrorLevel
            return

        ReceivePrompt := "Enter the receiving party:"
        InputBox, Receiver, Enter Receiver, %ReceivePrompt%, , , , , , , ,%Receiver%
        if ErrorLevel
            return
    }

    Quantity = %Qual%

    WinWaitActive, Trio SCS - Acctivate
    
    Send {Ctrl Down}{Shift Down}d{Shift Up}{Ctrl Up}
    Sleep, 2000
    if (printFlag) {
        PrintDoc()
        Sleep, 500
    }
    Sleep, 400
    Send {Enter}
    Sleep, 500
    ExportDoc()
    Sleep, 600
    Send {Enter}
    Sleep, 800
    Send {Enter}
    Sleep, 3000
    Send %POnum%-packslip-%Receiver%-%Qual%pcs-%PartNo%
    Sleep, 500
    Send {Alt Down}d{Alt Up}
    Sleep, 500
    Send P:\Warehouse\Jason backup\Migration\Inventory Upload
    Sleep, 200
    Send {Enter}
    Sleep, 250
    Send {Alt Down}s{Alt Up}
    Sleep, 200
    CloseReportWin()
    Sleep, 800

    Send {Ctrl Down}{Shift Down}t{Shift Up}{Ctrl Up}
    Sleep, 2000
    ExportDoc()
    Sleep, 150
    Send {Enter}
    Sleep, 50
    Send {Enter}
    Sleep, 1000
    Send %POnum%-serials-%Receiver%-%Qual%pcs-%PartNo%
    Sleep, 120
    Send {Alt Down}d{Alt Up}
    Sleep, 50
    Send P:\Warehouse\Jason backup\Migration\Inventory Upload
    Sleep, 50
    Send {Enter}
    Sleep, 150
    Send {Alt Down}s{Alt Up}
    Sleep, 200
    CloseReportWin()

    Send {Ctrl Down}{Shift Down}f{Shift Up}{Ctrl Up}
    Sleep, 4000
    UploadButton()
    Sleep, 3000
    Send {Alt Down}d{Alt Up}
    Sleep, 200
    Send P:\Warehouse\Jason backup\Migration\Inventory Upload
    Sleep, 200
    Send {Enter}
    Sleep, 300
    Send {Alt Down}n{Alt Up}
    Sleep, 400
    Send "%POnum%-packslip-%Receiver%-%Qual%pcs-%PartNo%" "%POnum%-serials-%Receiver%-%Qual%pcs-%PartNo%"
    Sleep, 200
    Send {Alt Down}o{Alt Up}
    Sleep, 300
    Send {Enter}
}

ObaDump() {
    WinActivate, rawdump - Excel
    Send %POnum%{Right}
    Sleep, 30
    Send %Quantity%{Right}
    Sleep, 30
    Send %PartNo%
    Sleep, 30
    Send {Down}{Home}
}

    ^!g::
    IfWinNotExist, rawdump - Excel
        RunWait, excel.exe "P:\Warehouse\Jason backup\Migration\rawdump.xlsx"
        WinActivate, Trio SCS - Acctivate
    PrepOutbound()
    ObaDump()
    return

VPDCheck() {
    Sleep, 25
    Send {Alt Down}d{Alt Up}
    Sleep, 100
    Send c{Right}
    Sleep, 100
    Send i
    VPDPrompt := "Incorrect VPD? enter y to exit."
    InputBox, Vcheck, Enter Vcheck, %VPDPrompt%
    if ErrorLevel
        return
    if (Vcheck = "y") {
        return
    }
    Send {Space}
    Sleep, 100
    Send {tab}{tab}
    Sleep, 100
    Send {Down}
    Sleep, 100
    VPDCheck()
}

    ^!m::
    VPDCheck()
    return

    ^!5::
    Send Maro2268123{!}{Enter}
    return

    ^!h::
    MsgBox,
    (
        Available Commands (all are togged by Ctrl+Alt+<key>):
        a: new note
        w: PO parts received note to buyer and Anne (opt.)
        q: All parts received to Anne
        e: Product Arrived note to buyer
        s: Shipped and invoice note to Anne/Julio
        c: Copy-Paste from Excel
        d: auto-ship notes from excel (see doc.)
        t: tracking note
        k: serial entry toggle
        f: RMA note to Julio
        g: Prep outbound 
        h: help
        u: AutoCopy
        m: VPD check

        1: set completion flag
        2: set date of entry
        3: toggle printing
        4: toggle Tested flag
        5: SV password
    )
    ;, ,Receiving Help
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