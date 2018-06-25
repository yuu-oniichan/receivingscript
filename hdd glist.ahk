^!s::
	QtyPrompt := "Enter number of items"
    InputBox, Qual, Enter Qual, %QtyPrompt%
    if ErrorLevel
        return
	
    Loop, %Qual% {
    	Send {Tab}{Tab}
    	Sleep, 150
    	Send {Space}
    	Sleep, 150
    	Send f{Down}
    	Sleep, 150
    	Send {Enter}
    	Sleep, 250
    	Send {Tab}
        Sleep, 150
        Send {Tab}
    	Sleep, 150
    	Send {Enter}
    	Sleep, 1200
    	Send {Enter}
    	Sleep, 600
    	Send {Tab}
    	Sleep, 200
    	Send {Down}
    	Sleep, 80
    	Send {Down}
    	Sleep, 80
    	Send {Down}
    	Sleep, 80
    	Send {Down}
    	Sleep, 120
    	Send {Enter}
    	Sleep, 200
    	Send {Enter}
    	Sleep, 150
    	Send {Alt Down}{F4}{Alt Up}
    	Sleep, 150
    	Send {Alt Down}{F4}{Alt Up}
    	Sleep, 150
    	Send {Tab}
    	Sleep, 150
    	Send {Down}
    	Sleep, 800
    }

	return	