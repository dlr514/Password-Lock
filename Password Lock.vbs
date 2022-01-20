Set objShell = CreateObject("Wscript.Shell")
dim password
password=InputBox("Please Enter Password:","3 - Tries Left")
if password = ("---------YOUR PASSWORD HERE---------") then
	dim correct
	correct =MsgBox("Correct Password!",64,"correct")
	objShell.Run("-------YOUR LINK HERE--------")
Else
	dim again
	again =MsgBox("Incorect Password! Do You Want To Try Again?",53,"Incorect Password!")
	If again = 4 Then
	dim password2
	password2=InputBox("Please Enter Password:","2 - Tries Left")
	if password2 = ("---------YOUR PASSWORD HERE---------") then
		dim correct2
		correct2 =MsgBox("Correct Password!",64,"correct")
		objShell.Run("-------YOUR LINK HERE--------")
	Else
		dim again2
		again2 =MsgBox("Incorect Password! Do You Want To Try Again?",53,"Incorect Password!")
		If again2 = 4 Then
		dim password3
		password3=InputBox("Please Enter Password:","1 - Tries Left")
		if password3 = ("---------YOUR PASSWORD HERE---------") then
			dim correct3
			correct3 =MsgBox("Correct Password!",64,"correct")
			objShell.Run("-------YOUR LINK HERE--------")
		Else
			dim again3
			again3 =MsgBox("Incorect Password! Do You Want To Try Again?",53,"Incorect Password!")
			If again3 = 4 Then
				dim incorect
						incorect =MsgBox("Too many incorect passwords! The program will now lock!",16,"WARNIG!!")
				objShell.Run("-------YOUR FAIL LINK HERE--------")
					end if
				end if
			end if
		end if
	end if
end if
