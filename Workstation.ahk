#NoEnv	; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn	 ; Enable warnings to assist with detecting common errors.
sendmode Input	; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%	 ; Ensures a consistent starting directory.
DetectHiddenText, On
DetectHiddenWindows, On
#Hotstring EndChars `t

;GuiSize:
		;GuiWidth = %A_GuiWidth%
		;GuiHeight = %A_GuiHeight%
		;ScreenWidth = %A_ScreenWidth%
		;ScreenHeight = %A_ScreenHeight%
		;MsgBox % A_ScreenHeight . " " . GuiHeight . " - " . y
		Return
		
		
+!r::
	Reload
	Return


;; SET UP WORKSTATION ;;

IfWinNotExist, Novell GroupWise - Jonathan Powers
{
	MsgBox, 1, Launch, Launch Default Workstation?
	IfMsgBox, Ok 
	{
		Run, "C:\Program Files\Mozilla Firefox\firefox.exe"
		Run, "C:\Program Files\Novell\GroupWise\grpwise.exe"
		Run, "C:\Users\Public\Desktop\cpr+.lnk"
		Run, "H:\Miscellaneous\products.xlsx"
		return
	}
}

;; CPR PLUS ;;

#IfWinActive, ahk_exe cpr.exe

+!n::
{
	MsgBox, 3, New Patient, Create New Patient?`nPress 'No' to Load Existing
	IfMsgBox, Yes
	{
		ClearPatient()
		return
	}
	IfMsgBox, No
	{
		Load()
		Return
	}
	IfMsgBox, Cancel
	{
		Dashboard()
		return
	}
}

ClearPatient()
{
	global 
	
	uid += 1
	ln := ""
	fn := ""
	dob := ""
	zip := ""
	addr := ""
	ph := ""
	mf := ""
	dos := ""
	iid := ""
	sn := ""
	npi := ""
	elig := ""
	auth := ""
	ins := ""
	alt := ""
	altnotes := ""
	altelig := ""
	notes := ""
	created := A_Now
	
	Insurance()
	Return
}

Insurance()
{
	global 

	Gui, New, AlwaysOnTop
	Gui, Font, s14, Helvetica
	Gui, Add, Button,, &HNJH
	Gui, Add, Button,, &BLUE
	Gui, Add, Button,, &UHC/AMC
	Gui, Add, Button,, &IBC
	Gui, Add, Button,, &AMH
	Gui, Add, Button,, AM&G
	Gui, Add, Button,, A&ET
	Gui, Add, Button,, &CIG
	Gui, Add, Button,, &QUAL
	Gui, Add, Button,, &MCAID
	Gui, Show, AutoSize
	;MsgBox, %GuiWidth% . " " . %GuiHeight%
	return

   ButtonHNJH:
   {
		elig := "https://navinet.navimedix.com/insurers/hnjh?start"
		ins := "Horizon NJ Health DME"
		notes := "Horizon NJ Health`n800-682-9091`nNo Auth Required"
		Gui, Submit
		Dash1()
		return
   }
   ButtonBLUE:
   {
		elig := "https://navinet.navimedix.com/insurers/horizon?start"
		ins := "Carecentrix Horizon DME"
		notes := "In State`nHorizon BCBS`n1-800-355-2583`nAuth through CareCentrix"
		alert := "Three Letter Prefix"
		auth := "https://www.carecentrixportal.com/ProviderPortal/homePage.do"
		
		alt := "Carecentrix Blue Card DME"
		altnotes := "Out of State`nCall for Eligibility`n1-800-676-2583`nNJ Residents:`nAuth through Carecentrix`nNon-Residents:`nAuth through Home Plan"
		Gui, Submit
		Dash1()
		return
   }
   ButtonUHC/AMC:
   {
		elig := "https://ptpsr.uhc.com/prweb/CRExternal"
		ins := "United Healthcare"
		notes := "United Healthcare`n877-842-3210`nNo Auth Required`nInsurance ID Starts with 8 or 9"

		alt := "AMERICHOICE DME II"
		altnotes := "UHCCP/Americhoice`n1-888-362-3368`nNo Auth Required`nInsurance ID Starts with 1"
		altalert := "MUST USE NBN 03"
		Gui, Submit
		Dash1()	  
		return
   }
   ButtonIBC:
   {
		elig := "https://navinet.navimedix.com/insurers/ibc?start"
		ins := "IBC PERSONAL CHOICE DME"
		notes := "Independence Blue Cross`n800-275-2583`No Auth Required"
		Gui, Submit
		Dash1()	  
		return
   }
   ButtonAMH:
   {
		elig := "https://navinet.navimedix.com/insurers/amerihealth-nj?start"
		ins := "AMERIHEALTH - New Jersey DME"
		notes := "AmeriHealth`n800-356-7899"
		altnotes := "AmeriHealth Administrators`n800-841-5328`No Auth Required"
		alt := "AMERIHEALTH  ADMINISTRATORS DME"
		altelig := "https://navinet.navimedix.com/insurers/aha-tpa?start"
		Gui, Submit
		Dash1()	  
		return
   }
   ButtonAMG:
   {
		elig := ""
		ins := "AMERIGROUP - NJ MEDICAID DME"
		notes := "Amerigroup`n1-800-454-3730`nMust Call for Eligibility`nNo Auth Required"
		Gui, Submit
		Dash1()	  
		return
   }
   ButtonAET:
   {
		elig := "https://navinet.navimedix.com/insurers/aetna?start"
		ins := "AETNA"
		notes := "Aetna`n800-624-0756`nNo Auth Required"
		Gui, Submit
		Dash1()	  
		return
   }
   ButtonCIG:
   {
		elig := "https://navinet.navimedix.com/insurers/cigna?start"
		ins := "GENTIVA CARECENTRIX DME"
		notes := "Cigna`800-794-7882`nAuth through CareCentrix as 'Cigna PPO'"
		alert := "Auth Cigna PPO"
		auth := "https://www.carecentrixportal.com/ProviderPortal/homePage.do"
		Gui, Submit
		Dash1()	  
		return
   }
   ButtonQUAL:
   {
		elig := "https://navinet.navimedix.com/insurers/qualcare?start"
		ins := "Qualcare"
		notes := "QualCare`n800-992-6613`No Auth Required"
		Gui, Submit
		Dash1()
		return
	}
	ButtonMCAID:
   {
		elig := "https://www.njmmis.com/mevs.aspx"
		ins := "MEDICAID DME"
		notes := "Medicaid`nCall for Eligibility`n800-776-6334`nNo Auth Required"
		alert := "MUST HAVE SCRIPT"
		Gui, Submit
		Dash1()
		return
	}
}

InsLabel()
{
	global
	
	if (!ins){
		return "None"
	}
	if (ins = "MEDICAID DME"){
		return "Medicaid"
	}
	if (ins = "Qualcare"){
		return "Qualcare"
	}
	if (ins = "GENTIVA CARECENTRIX DME"){
		return "Cigna"
	}
	if (ins = "AETNA"){
		return "Aetna"
	}
	if (ins = "AMERIGROUP - NJ MEDICAID DME"){
		return "Amerigroup"
	}
	if (ins = "AMERIHEALTH - New Jersey DME"){
		return "Amerihealth"
	}
	if (ins = "AMERIHEALTH  ADMINISTRATORS DME"){
		return "Amerihealth Admin"
	}
	if (ins = "IBC PERSONAL CHOICE DME"){
		return "IBC"
	}
	if (ins = "United Healthcare"){
		return "UnitedHealthCare"
	}
	if (ins = "AMERICHOICE DME II"){
		return "Americhoice"
	}
	if (ins = "Carecentrix Horizon DME"){
		return "Horizon Blue Card"
	}
	if (ins = "Horizon NJ Health DME"){
		return "Horizon NJ Health"
	}
	return "Other"
}

Dash1()
{
	Global
	
	Gui, New, AlwaysOnTop
	Gui, Font, s16, Helvetica
	Gui, Add, Text, Center, %notes%.`n`n%altnotes%
	
	if elig
		Gui, Add, Link,, Eligibility - <a href="%elig%">Click</a>
	if auth
		Gui, Add, Link,, Auth - <a href="%auth%">Click</a>`n
	if altelig
		Gui, Add, Link,, Alt Eligibility - <a href="%altelig%">Click</a>
	
	Gui, Add, Button,, Next
	if Alt
		Gui, Add, Button, x+m, Alt
	Gui, Show, AutoSize x1215 y10
	;MsgBox, %GuiWidth% . " " . %GuiHeight
	Gui, Show, AutoSize x1215 y10
	
	Return
	
	ButtonNext:
	{
		Gui, Submit
		Input()
		return
	}
	ButtonAlt:
	{
		ins := alt
		Gui, Submit
		Input()
		return
	}
}

Input()
{	
	global
   
	Gui, New, AlwaysOnTop
	Gui, Font, s13, Helvetica
   
	Gui, Add, Text, , Last Name
	Gui, Add, Edit, Uppercase vln, %ln%
	Gui, Add, Text, , First Name
	Gui, Add, Edit, Uppercase vfn, %fn%
	Gui, Add, Text, , Date of Birth
	Gui, Add, Edit, Uppercase vdob, %dob%
	Gui, Add, Text, , Zip Code
	Gui, Add, Edit, Uppercase vzip, %zip%
   
	Gui, Add, Text, , `nAddress
	Gui, Add, Edit, Uppercase vaddr, %addr%
	Gui, Add, Text, , Phone Number
	Gui, Add, Edit, Uppercase vph, %ph%
	Gui, Add, Text, , Gender
	Gui, Add, Edit, Uppercase vmf, %mf%
	Gui, Add, Text, , Date of Service
	Gui, Add, Edit, Uppercase vdos, %dos%
   
	Gui, Add, Text, , `nInsurance ID
	Gui, Add, Edit, Uppercase viid, %iid%
	Gui, Add, Text, , NPI
	Gui, Add, Edit, Uppercase vnpi, %npi%
	Gui, Add, Text, , Serial Number
	Gui, Add, Edit, Uppercase vsn, %sn%
   
	Gui, Add, Button, Default gOK, OK
	
	Gui, Show, Autosize x1250 y15
	;x := ScreenWidth - GuiWidth - 15
	;y := ScreenHeight - GuiHeight - 95
	;MsgBox % GuiWidth . " " . GuiHeight
	;Gui, Show, x%x% y%y%
	
	
	return

   OK:
	{
		Gui, Submit
		clipboard := dob
		Dashboard() 
		return
	}
}

Dashboard()
{
	global
	
	insure := InsLabel()

	Gui, New, AlwaysOnTop
	Gui, Font, s13, Helvetica
	Gui, Add, Text, Center, Press Button to Copy Value`nOr type [Shortcut] and hit Tab`n
	
	if auth
		Gui, Add, Link, xm, Auth - <a href="%auth%">Click</a>
	if alert
		Gui, Add, Text, xm, %alert%
	if alt
		Gui, Add, Text, xm, %altalert%
   
	Gui, Add, Button, xm, Last Name [ln]
	Gui, Add, Text, x+m yp+7, %ln%
	Gui, Add, Button, xm, First Name [fn]
	Gui, Add, Text, x+m yp+7, %fn%
	Gui, Add, Button, xm, Date of Birth [dob]
	Gui, Add, Text, x+m yp+7, %dob%
	Gui, Add, Button, xm, Zip Code [zip]
	Gui, Add, Text, x+m yp+7, %zip%`n`n
   
	Gui, Add, Button, xm, Address [addr]
	Gui, Add, Text, x+m yp+7, %addr%
	Gui, Add, Button, xm, Phone Number [ph]
	Gui, Add, Text, x+m yp+7, %ph%
	Gui, Add, Button, xm, Gender [mf]
	Gui, Add, Text, x+m yp+7, %mf%
	Gui, Add, Button, xm, Date of Service [dos]
	Gui, Add, Text, x+m yp+7, %dos%`n`n
   
	Gui, Add, Button, xm, Insurance ID [iid]
	Gui, Add, Text, x+m yp+7, %iid%
	Gui, Add, Button, xm, NPI [npi]
	Gui, Add, Text, x+m yp+7, %npi%
	Gui, Add, Button, xm, Insurance [ins]
	Gui, Add, Text, x+m yp+7, %insure%
	Gui, Add, Button, xm, Serial Number [sn]
	Gui, Add, Text, x+m yp+7, %sn%`n`n
	Gui, Add, Button, xm, Save
	Gui, Add, Button, x+m, Update Values 
	
	Gui, Show, Autosize x1250 y 15

	;Gui, Show, % "x" A_ScreenWidth - A_GuiWidth " y" A_ScreenHeight - A_GuiHeight " w" A_GuiWidth " h" A_GuiHeight
	
	;MsgBox % x . " " . y . " " . w . " " . h
	;Gui, Show, x%A_GuiWidth% y%A_GuiHeight%
	return
	
	ButtonLastName[ln]:
	{
		Clipboard = %ln%
		Return
	}
	ButtonFirstName[fn]:
	{
		Clipboard = %fn%
		Return
	}
	ButtonDateofBirth[dob]:
    {
		clipboard = %dob%
		return
    }
	ButtonZipCode[zip]:
	{
		Clipboard = %zip%
		Return
	}
	ButtonAddress[addr]:
	{
		clipboard = %addr%
		return
	}
	ButtonGender[mf]:
	{
		clipboard = %mf%
		return
	}
	ButtonDateofService[dos]:
	{
		clipboard = %dos%
		return
	}
	ButtonPhoneNumber[ph]:
	{
		clipboard = %ph%
		return
	}
	ButtonInsuranceID[iid]:
	{
		clipboard = %iid%
		return
	}
	ButtonSerialNumber[sn]:
	{
		clipboard = %sn%
		return
	}
	ButtonNPI[npi]:
	{
		clipboard = %npi%
		return
	}
	ButtonSave:
	{
		Gui, Submit
		Save()
		Return
	}
	ButtonUpdateValues:
	{
		Gui, Submit
		Input()
		Return
	}
}

#IfWinActive, Workstation.ahk
!s::
{
	Gui, Submit
	Save()
	return
}
#IfWinActive

Save()
{	
	Global
	
	record := Array(uid, created, A_Now, ln, fn, dob, zip, addr, ph, mf, dos, iid, sn, npi, ins)
	saveFile := "`n"
	
	for index, value in record{
		saveFile := saveFile . value . ","
	}
	
	try
	{	
		file := FileOpen("../billing.csv", "a")
		;MsgBox, File Loaded
		file.write(saveFile)
		MsgBox, Successful Save!
		file.close()
	} catch e{
		MsgBox, Failed to Save!
	}
	
	MsgBox, 3, New Patient?, Create New Patient Now?
	IfMsgBox, Yes
	{
		Gui, Submit
		ClearPatient()
		Return
	}
	IfMsgBox, No
	{
		Gui, Submit
		Dashboard()
		Return
	}
	IfMsgBox, Cancel
	{
		Gui, Submit
		Return
	}
	
	Return
}

Load()
{
	Global
	
	Gui, New, AlwaysOnTop
	Gui, Add, Text,, What would you like to search by?
	Gui, Add, Button, Default, &Last Name
	Gui, Add, Button,, &Date of Service
	Gui, Add, Button,, Insurance &ID
	Gui, Show
	Return
	
	ButtonLastName:
	{
		Gui, Submit
		index = 4
		Read()
		Return
	}
	ButtonDateofService:
	{
		Gui, Submit
		index = 11
		Read()
		Return
	}
	ButtonInsuranceID:
	{
		Gui, Submit
		index = 12
		Read()
		Return
	}
}

Read()
{	
	Global

	Gui, New, AlwaysOntop
	Gui, Add, Text,, Enter Search Term
	Gui, Add, Edit, vsearch
	Gui, Add, Button, Default, Ok
	Gui, Show
	Return
	

	ButtonOK:
	{	
		Gui, Submit
	}
	
	Loop, Read, ../billing.csv
	{
		StringReplace, Cleaned, A_LoopReadLine, `%, , All
		Record := StrSplit(Cleaned,",") 
		parse := Parse(Record)
		
		if InStr(Record[index], search)
		{
			Gui, New, AlwaysOntop
			Gui, Add, Text,, Is This the Record?
			Gui, Add, Text,, %parse%
			Gui, Add, Button, Default, &Yes
			Gui, Add, Button,, &No
			Gui, Show
			Return
			
			ButtonYes:
			{
				Gui, Submit
				LoadPatient(Record)
				Return
			}
			ButtonNo:
			{	
				Gui, Submit
				Continue
			}			
		}
	}
	MsgBox, Not Found, Trying Again
	Gui, Submit
	Load()
	Return
}

Parse(Record)
{	
	FormatTime, Date, Record[3], LongDate
	FormatTime, Time, Record[3], h:m:ss tt
	FormatTime, Seen, Record[11], ShortDate
	string := "Patient No " . Record[1] . ": " Record[4] . " " . Record[5] . ", seen " . Seen . ".`nLast Edited" . Date . " at " . Time
	return string
}

LoadPatient(LoadFile)
{
	Global
	
	uid := LoadFile[1]
	created := LoadFile[2]
	ln := LoadFile[4]
	fn := LoadFile[5]
	dob := LoadFile[6]
	zip := LoadFile[7]
	addr := LoadFile[8]
	ph := LoadFile[9]
	mf := LoadFile[10]
	dos := LoadFile[11]
	iid := LoadFile[12]
	sn := LoadFile[13]
	npi := LoadFile[14]
	ins := LoadFile[15]
	
	Dashboard()
	Return
}

; HOTKEYS

+!^i::
{
	PostMessage, 0x001C, 003305D4, 0, , CPR+ Equipment Manager

}

+!i:: ;INVENTORY CHECK
{	
	;sn := "ad16I14124"
	;SetMouseDelay, 250
	WinActivate, ahk_exe cpr.exe
	WinActivate, Main Menu
	Sleep 600
	Loop
	{
		IfWinActive, Main Menu
			SendInput 6
			Sleep, 600
			Break
		Sleep, 200
	}
	Sleep, 600
	Loop
	{
		IfWinActive, Inventory Menu
			SendInput 9
			Sleep, 600
			Break
		Sleep, 200
	}
	WinActivate, CPR+ Equipment Manager
	Sleep, 600
	Loop
	{	
		IfWinActive, CPR+ Equipment Manager
			SendInput {F6}
			Sleep, 600
			Break
		Sleep, 200
	}
	Loop
	{
		IfWinActive, Equipment Manager - Rental search
			Sleep, 1200
			SendInput %sn%
			Sleep, 1200
			Break
		Sleep, 200
	}
	Send {F6}
	return
}
;


; HOTSTRINGS
::ln::
{
	if %ln%
	{
		SendInput %ln%
		SendInput {Tab}
	}Else
	{
		Send ln
	}
	Return
}
::fn::
{
	if %fn%
	{
		SendInput %fn%
		SendInput {Tab}
	}
	Else
	{
		Send fn
	}
	Return
}
::zip::
{
	if %zip%
	{
		SendInput %zip%	
	}
	Else
	{
		Send zip
	}
	Return
}
::sn::
{
	if %sn%
	{
		SendInput %sn%
	}
	Else
	{
		Send sn
	}
	Return
}
::dos:: 
{
	if %dos%
	{
		#IfWinActive, ahk_exe cpr.exe
		Send {F11}
		#IfWinActive
		SendInput %dos%
		Send {Enter}
	}
	Else
	{
		Send dos
	}
	return
}
::dob:: 
{
	if %dob%
	{
		#IfWinActive, ahk_exe cpr.exe
		Send {F11}
		#IfWinActive
		SendInput %dob%
		Send {Enter}
	}
	Else
	{
		Send dob
	}
	return
}
::addr::
{
	if %addr%
	{
		SendInput %addr%
		SendInput {Enter}
	}
	Else
	{
		Send addr
	}
	return
}
::ph::
{	
	if %ph%
	{
		#IfWinActive, ahk_exe cpr.exe
		Send {F11}
		#IfWinActive
		SendInput %ph%
	}
	Else
	{
		Send ph
	}
	return
}
::npi::
{
	if %npi%
	{
		SendInput %npi%
	}
	Else
	{
		Send npi
	}
	return
}
::iid::
{
	if %iid%
	{
		SendInput %iid%
	}
	Else
	{
		Send iid
	}
	return
}
::ins::
{
	if %ins%
	{
		SendInput %ins%
	}
	Else
	{
		Send ins
	}
	return
}
;


; FORM FILLERS
;::demo:: 
{
   SendInput %addr%
   Sleep, 500
   Send {Tab 2}
   Sleep, 500
   Send {Tab 3}
   Sleep, 1000	 
   Send {F11}
   Sleep, 500
   SendInput %ph%
   Sleep, 1000
   Send {Tab 4}
   Sleep, 500
   SendInput mf
   Sleep, 1000
   Send {Tab 3}
   Sleep, 500
   Send {F11}
   Sleep, 500
   SendInput %dos%
   Sleep, 1000
   Send {Tab 3}
   Sleep, 500
   Send {Tab 4}
   Sleep, 500
   Send {F11}
   SendInput %dos%
   Sleep, 500
   Send {Tab 6}
   Sleep, 500
   Send {F10}
   Sleep, 500
   Send I
   Sleep, 1000
   Send {Enter}
   Sleep, 500
   Send {F10}
   Sleep, 500
   Send CON
   Sleep, 1000
   Send {Enter}
   Sleep, 500
   Send {F10}
   Sleep, 500
   Send M
   Sleep, 1000
   Send {Enter}
  return
}
::cmn:: 
{
	SendInput, %dos%
	Sleep, 500
	SendInput, {Enter}
	Sleep, 500
	Send {F2}
	Loop
	{
		IfWinActive, Document Tracking
			Break
		Sleep, 50
	}
	Send {F2}
	Loop
	{
		IfWinActive, Inscrybe Interface
			Break
		Sleep, 50
	}
	Send n
  return
}
;
#IfWinActive

;;
;; End CPR Script
;;


;; GROUPWISE ;;

; Main Window
#IfWinActive, Novell GroupWise - Jonathan Powers Home

^t:: WinMenuSelectItem, Novell GroupWise - Jonathan Powers Home, , File, New, Tasklist Item

#IfWinActive

; TodoView
#IfWinActive, ahk_class OFTodoView

^Enter:: Click 45, 70 

#IfWinActive


;; ACROBAT - FAX ;;
+!F::
{
	Run, "H:\Forms\FaxSheet_Infusions.pdf"
	return
}

#IfWinActive, ahk_class AcrobatSDIWindow

::diag::
(
Attached consignment form is missing a diagnosis code.
)

::sig::
(
Please return the complete form so it can be submitted to the insurance companies.

Questions, please call 856.669.0211. My extension is 200224.

Thank you.
)

#IfWinActive
; END ACROBAT


;; GLOBALS ;;
::phy::Physician

^`:: ExitApp