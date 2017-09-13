;~ This library is the Product of tank
;~ based on COM.ahk from Sean http://www.autohotkey.com/forum/topic22923.html
;~ standard library is the work of tank and updates may be downloaded from
;~ http://www.autohotkey.net/~tank/iWeb.zip
;~ http://www.autohotkey.com/forum/viewtopic.php?t=51270
;~ complete API
/*
iWeb_Init()
iWeb_Term()
iWeb_NewIe()
iWeb_Model(h=550,w=900)
iWeb_GetWin(Name="")
iWeb_GetIE(Name="A")
iWeb_GetDocument(title="A")
iWeb_BRSRfromOBJ(obj)
iWeb_Release(pdsp)
iWeb_Nav(pwb,url)
iWeb_complete(pwb)
iWeb_DomWin(pdsp,frm="")
iWeb_inpt(i)
iWeb_getDomObj(pwb,obj,frm="")
iWeb_setDomObj(pwb,obj,t,frm="")
iWeb_FindbyText(needle,win="A",property="",offset=0,frm="")
iWeb_Checked(pwb,obj,checked=1,sIndex=0,frm="")
iWeb_SelectOption(pdsp,sName,selected,method="selectedIndex",frm="")
iWeb_TableParse(pdsp,table,row,cell,frm="")
iWeb_GetElementByAll(pdsp,obj,index=0,frm="")
iWeb_GetElementsByTag(pdsp,obj,index=0,frm="")
iWeb_Offset(pdsp,offset=0)
iWeb_FireEvents(ele)
iWeb_Attributes(element)
iWeb_TableLength(pdsp,TableRows="",TableRowsCells="",frm="")
iWeb_clickDomObj(pwb,obj,frm="")
iWeb_clickText(pwb,t,frm="")
iWeb_clickHref(pwb,t,frm="")
iWeb_clickValue(pwb,t,frm="")
iWeb_execScript(pwb,js,frm="")
iWeb_getVar(pwb,var,frm="")
iWeb_escape_text(txt)
iWeb_striphtml(HTML)
iWeb_UrlEncode( String )
iWeb_uriDecode(str)
iWeb_Txt2Doc(t)
iWeb_Activate(sTitle)
*/




;~ getting/destroying browser handles*

	;~ A new internet explorer window
	iWeb_NewIe()
		{
		Return	pwb := (pwb := ComObjCreate("InternetExplorer.Application") ) ? pwb.Visible:=True" : 0
		}
	;~ New internet explorer window always on top with titlebar only
	iWeb_Model(h=550,w=900)
		{
		If	pwb := (pwb := iWeb_newIe()) ? (pwb, pwb.MenuBar:=0, pwb.ToolBar:=0, pwb.Resizablev0, pwb.AddressBar:=0, pwb.StatusBar:=0, pwb.Height:=h, pwb.Width:=w) : 
		WinSet,AlwaysOnTop,On,% "ahk_ID " pwb.hwnd
		Return	pwb
		}
	;~ reuse an existing tab or window
	iWeb_GetWin( ByRef sTitle = "", ByRef iHWND = "", ByRef sURL = "", ByRef sHTML = "" )
		{		
		;; find where windows believes IE is installed
		;; certain corp installs may have this in other than expected folders
		IE_path := IE_GetPath()
		
		;~ this function is pointless if no instance of IE is open
		;~ one edit you mihgt make is to have this function open IE and maybe go to the home page
		if ( !winexist( "ahk_class IEFrame" ) )
			{
			MsgBox, 4112, NO IE Window Found, The Macro will end
			ExitApp
			}
		
		if sTitle
			clean_IE_Title( sTitle ) 
		;; ok this function should look at all the existing IE instances and build a reference object
		; List all open Explorer and Internet Explorer windows:
		oIE := Object()
		matches := 0
		;~ msgbox % sTitle
		;~ msgbox % iHWND
		;~ msgbox % sURL
		for window,k in ComObjCreate("Shell.Application").Windows
			if ( "Internet Explorer" = window.Name)
				{
				possiblematch := true
				if !window.document
					continue
				pdoc := IHTMLWindow2_from_IWebDOCUMENT( window.document ).document
				if ( possiblematch && sTitle && !instr( pdoc.title, sTitle ) )
					possiblematch := false
				
				if ( possiblematch && sHTML && !instr( pdoc.documentelement.outerhtml, sHTML ) )
					possiblematch := false
				
				if ( possiblematch && sURL && !instr( pdoc.url, sURL ) )
					possiblematch := false
				
				if ( possiblematch && iHWND > 0 && window.HWND != iHWND )
					possiblematch := false		
				;~ MsgBox % sTitle := pdoc.title
				if ( possiblematch )
					{
					;~ windowsList .= k " => " ( clipboard := window.FullName ) " :: " pdoc.title " :: " pdoc.url "`n"
					matches++
					sTitle := pdoc.title
					sURL := pdoc.url
					iHWND := window.HWND
					sHTML := pdoc.documentelement.outerhtml
					oIE := window
					}
				ObjRelease( pdoc )
				}
				
		if ( matches > 1 )
			{
			MsgBox, 4112, Too many Matches ,  Please modify your criteria or close some tabs/windows and retry
			ExitApp
			}
		;~ msgbox % sHTML
		return oIE
		}
;~ 	Usefull for particularly stubborn windows but be aware it can fail you if the page navigates and doesnt give you access to browser objects and properties
	iWeb_GetIE(Name="A") 
		{               ;// based on ComObjQuery docs
		static msg := DllCall( "RegisterWindowMessage", "str", "WM_HTML_GETOBJECT" )
			, IID_IWebDOCUMENT := "{332C4425-26CB-11D0-B483-00C04FD90119}"
		
		SendMessage msg, 0, 0, Internet Explorer_Server1, %Name%
		
		if (ErrorLevel != "FAIL") 
			{
			lResult := ErrorLevel
			VarSetCapacity( GUID, 16, 0 )
			if DllCall( "ole32\CLSIDFromString", "wstr", IID_IWebDOCUMENT, "ptr", &GUID ) >= 0 
				{
				DllCall( "oleacc\ObjectFromLresult", "ptr", lResult, "ptr", &GUID, "ptr", 0, "ptr*", IWebDOCUMENT )
				return  IWebBrowserApp_from_IWebDOCUMENT( IWebDOCUMENT )
				}
			}
		}
	IE_GetPath(){
		Static IE_path
		;; find where windows believes IE is installed
		;; certain corp installs may have this in other than expected folders
		if !IE_path
			RegRead, IE_path, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\IEXPLORE.EXE
		;~ MsgBox % IE_path
		;; Perhaps policies prevent reading this key
		if ( ErrorLevel || !IE_path )
			IE_path := "C:\Program Files\Internet Explorer\iexplore.exe"
		
		;; make sure it installed
		if !FileExist( IE_path )
			{
			MsgBox, 4112, Internet Explorer Not Found, IE does not appear to be installed`nCannot continue `nClick OK to Exit!!!
			ExitApp
			}
		
		return IE_path
		}
	IE_GetVersion(){
		FileGetVersion, version, % IE_GetPath()
		if ErrorLevel = 1
			ieVer = 11
		else
			ieVer := SubStr(version, 1, InStr(version, ".")-1)
		
		return ieVer
		}
	clean_IE_Title( ByRef sTitle = "" ) 
		{
		return sTitle := RegExReplace( sTitle ? sTitle : active_IE_Title(), IE_Suffix() "$", "" )
		}

	IE_Suffix() 
		{
		static sIE_Suffix
		if !sIE_Suffix
			{
			;; HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main
			RegRead, sIE_Suffix, HKCU, Software\Microsoft\Internet Explorer\Main, Window Title ;, Windows Internet Explorer,
			sIE_Suffix := " - " sIE_Suffix
			}
		return sIE_Suffix
		}

	active_IE_Title() ;; returns the title of the topmost browser if exists from the stack
		{
		sTitle := "NO IE Window Open"
		if winexist( "ahk_class IEFrame" )
			{
			titlematchMode := A_TitleMatchMode
			titlematchSpeed := A_TitleMatchModeSpeed
			SetTitleMatchMode, 2	
			SetTitleMatchMode, Slow
			WinGetTitle, sTitle, %sIE_Suffix% ahk_class IEFrame
			SetTitleMatchMode, %titlematchMode%	
			SetTitleMatchMode, %titlematchSpeed%
			}
		return RegExReplace( sTitle, IE_Suffix() "$", "" )
		}
		
		
		
	IHTMLWindow2_from_IWebDOCUMENT( IWebDOCUMENT )
		{
		static IID_IHTMLWindow2 := "{332C4427-26CB-11D0-B483-00C04FD90119}"  ; IID_IHTMLWindow2
		return ComObj(9,ComObjQuery( IWebDOCUMENT, IID_IHTMLWindow2, IID_IHTMLWindow2),1)
		}

	IWebDOCUMENT_from_IWebDOCUMENT( IWebDOCUMENT ) ;bypasses certain security issues
		{
		return IHTMLWindow2_from_IWebDOCUMENT( IWebDOCUMENT ).document
		}

	IWebBrowserApp_from_IWebDOCUMENT( IWebDOCUMENT )
		{
		static IID_IWebBrowserApp := "{0002DF05-0000-0000-C000-000000000046}"  ; IID_IWebBrowserApp
		return ComObj(9,ComObjQuery( IHTMLWindow2_from_IWebDOCUMENT( IWebDOCUMENT ), IID_IWebBrowserApp, IID_IWebBrowserApp),1)
		}
	;~ ControlGet, hwnd, Hwnd,, Internet Explorer_Server1, ahk_class IEFrame
	IWebBrowserApp_from_Internet_Explorer_Server_HWND( hwnd, Svr#=1 ) 
		{               ;// based on ComObjQuery docs
		static msg := DllCall( "RegisterWindowMessage", "str", "WM_HTML_GETOBJECT" )
			, IID_IWebDOCUMENT := "{332C4425-26CB-11D0-B483-00C04FD90119}"
		
		SendMessage msg, 0, 0, Internet Explorer_Server%Svr#%, ahk_id %hwnd%
		
		if (ErrorLevel != "FAIL") 
			{
			lResult := ErrorLevel
			VarSetCapacity( GUID, 16, 0 )
			if DllCall( "ole32\CLSIDFromString", "wstr", IID_IWebDOCUMENT, "ptr", &GUID ) >= 0 
				{
				DllCall( "oleacc\ObjectFromLresult", "ptr", lResult, "ptr", &GUID, "ptr", 0, "ptr*", IWebDOCUMENT )
				return  IWebBrowserApp_from_IWebDOCUMENT( IWebDOCUMENT )
				}
			}
		}


;~ 	Usefull for particularly stubborn windows but be aware it can fail you if the page navigates and doesnt give you access to browser objects and properties
	iWeb_GetDocument(title="A")
		{
		Static
	;~ 	Compliments Sean taken nearly un modified from the IE Spy http://www.autohotkey.com/forum/viewtopic.php?t=48470
		
		static msg := DllCall( "RegisterWindowMessage", "str", "WM_HTML_GETOBJECT" )
			, IID_IWebDOCUMENT := "{332C4425-26CB-11D0-B483-00C04FD90119}"
		
		SendMessage msg, 0, 0, Internet Explorer_Server%Svr#%, ahk_id %hwnd%
		
		if (ErrorLevel != "FAIL") 
			{
			lResult := ErrorLevel
			VarSetCapacity( GUID, 16, 0 )
			if DllCall( "ole32\CLSIDFromString", "wstr", IID_IWebDOCUMENT, "ptr", &GUID ) >= 0 
				{
				DllCall( "oleacc\ObjectFromLresult", "ptr", lResult, "ptr", &GUID, "ptr", 0, "ptr*", IWebDOCUMENT )
				return  IWebDOCUMENT
				}
			}
			return false
		}
	
	iWeb_BRSRfromOBJ(obj)
	{ ;; accepts any child object and returns an instance of iexplore.exe
		static IID_IWebBrowserApp := "{0002DF05-0000-0000-C000-000000000046}"  ; IID_IWebBrowserApp
		return ComObj(9,obj, IID_IWebBrowserApp, IID_IWebBrowserApp),1)
	}

	;~ Navigate to a url
	iWeb_Nav(pwb,url)						; returns bool 
		{
		If  (!pwb || !url)		;	test to see if we have a valid interface pointer
			{
			MsgBox, 262160, Browser Navigation, The Browser you tried to Navigate to `n%url%`nwith is not valid
			Return						;	ExitApp if we dont
			}
	;~ 	
	;~ 	http://msdn.microsoft.com/en-us/library/aa752133(VS.85).aspx
		navTrustedForActiveX	=	0x0400
		pwb.Navigate(url,	navTrustedForActiveX,	"_self")
		iWeb_complete(pwb)
		Return							;	return the result(bool) of the complete function 
		}									;	nav function end
	;~ wait for a page to finish loading
	iWeb_complete(pwb)						;	returns bool for success or failure
	{	
		If  !pwb							;	test to see if we have a valid interface pointer
			sleep, 5000						;	ExitApp if we dont
		Else
		{
			loop 20							;	sets limit if itenerations to 40 seconds 80*500=40000=40 secs
				If not (rdy:=pwb.readyState = 4)
					Break				;	return success
				Else	Sleep,100					;	sleep .1 second between cycles
			loop 80							;	sets limit if itenerations to 40 seconds 80*500=40000=40 secs
				If (rdy:=pwb.readyState = 4)
					Break
				Else	Sleep,500					;	sleep half second between cycles
			Loop	80				
				If	(rdy:=pwbdocument.readystate="complete")
					Return 	1				;	return success
				Else	Sleep,100
		}
		Return 0						;	lets face it if it got this far it failed
	}								;	end complete
	;~ get the window onject from an object
	iWeb_DomWin(pdsp,frm="")
		{
		If	pWin	:=	IHTMLWindow2_from_IWebDOCUMENT( pdsp ) 
			{
			Loop, Parse, frm, `, 
				{
				frame:=pWin.document.all.item['" A_LoopField "'].contentwindow
				ObjRelease(pWin)
				pWin:=IHTMLWindow2_from_IWebDOCUMENT( frame )
				ObjRelease(frame)
				If	!pWin
					Return	False
				}
			Return	pWin
			}	
		Return False
		}
	;~ Determin if an element is a form input 
	iWeb_inpt(i)
		{
	;~ 			http://msdn.microsoft.com/en-us/library/ms534657(VS.85).aspx tagname property
		typ		:=	i.tagName
		inpt	:=	"BUTTON,INPUT,OPTION,SELECT,TEXTAREA" ; these things all have value attribute and is likely what i need instead of innerHTML
		Loop,Parse,inpt,`,
			if (typ	=	A_LoopField	?	1	:	"")
				Return 1
		Return
		}
;~ Functions that manipulate DOM

	iWeb_getDomObj(pwb,obj,frm="")
		{
		/*********************************************************************
		@depricated
		pwb	-	browser object
		obj	-	object reference; optionally, (name, id or index) of all value can be used
		frm -	frame reference; optionally, a comma delimited list of frames (name, id or index ) of all value can be used
		example of usage
		The below will try to get an object called 'username'
		iWeb_getDomObj(pwb,"username")
		*/
		If	itm		:=	iWeb_GetElementByAll(pwb,obj,0,frm)	;if this fails there really isnt any need to do below
			{
			T:=iWeb_inpt(itm) ? "value" : "innerHTML")
			rslt	.=	iWeb_uriDecode(itm.%T%) 
			iWeb_FireEvents(itm)
			ObjRelease(itm)
			}
		Return	rslt
		}

	iWeb_setDomObj(pwb,obj,t,frm="")
		{
		/*********************************************************************
		pwb	-	browser object
		obj	-	object reference; optionally 
		t	-	text to place in object; 
		frm -	frame reference; optionally, 
		Example Usage
		The below will take a browser object, try to get an object called 'username' and set its value/innerHTML to 'john'
		iWeb_setDomObj(pwb,"username","john")
		*/
		If	itm		:=	iWeb_GetElementByAll(pwb,obj,0,frm)	;if this fails there really isnt any need to do below
			{
			;~ 	http://www.autohotkey.com/forum/viewtopic.php?p=221631#221631 iWeb_uriDecode(str)
			v:=iWeb_inpt(itm) ? "Value" : "innerHTML"
			itm.%v% :=iWeb_UrlEncode(t)
			iWeb_FireEvents(itm)
			ObjRelease(itm)
			d=1
			}
		Return d
		}

	iWeb_FindbyText(needle,win="A",property="",offset=0,frm="")
		{
		If	pWin	:=	iWeb_DomWin(pwb,frm) 
			{
			If	oRange:=pWin.document.body.createTextRange
				{
				oRange.findText(needle)
				_res:=property ? pWin.Document.all.item[ oRange.parentElement.sourceIndex+offset].%property%) :  pWin.Document.all.item[ oRange.parentElement.sourceIndex+offset]
				ObjRelease(oRange)
				}	
			}
		Return	_res
		}




iWeb_GetElementByAll(pdsp,obj,sindex=0,frm=""){ ;; returns object
;~ 	COM_Error(0)
	Return element := (pWin	:=	iWeb_DomWin(pdsp,frm) ) ? (sindex > 0 ?  pWin.document.all.[obj].item",sindex : pWin.document.all.item[obj]),ObjRelease(pWin)) : 
}

;~ iWeb_GetElementsByTag(pdsp,tag,obj=0,frm=""){  ;; returns object
	;~ Return element := (pWin	:=	iWeb_DomWin(pdsp,frm) ) ? (COM_Invoke(pWin,"document.all.tags['" tag "'].item",obj),COM_Release(pWin)) : 
;~ }

;~ iWeb_Offset(pdsp,offset=0){
	;~ Return element := (pWin	:=	iWeb_DomWin(pdsp,frm) ) ? (COM_Invoke(pWin,"document.all",COM_Invoke(pdsp,"sourceIndex") + offset),COM_Release(pWin)) : 
;~ }

iWeb_FireEvents(ele)
{
	attributes:=iWeb_Attributes(ele)
	Loop,Parse,attributes, `n
		If	InStr(A_LoopField,"on")
			ele.%A_LoopField%
}



;~ functions that click 

	iWeb_clickDomObj(pwb,obj,frm="")
	{
	/*********************************************************************
	pwb	-	browser object
	obj	-	object reference; optionally, a comma delimited list of references (name, id or index) of all value can be used
	frm -	frame reference; optionally, a comma delimited list of frames (name, id or index ) of all value can be used
	Example of Usage
	The below will take a browser object and try to click an object called username
	iWeb_getDomObj(pwb,"username")
	This will cycle thru and attempt to click 3 separate objects (username, pass and 3) 
	iWeb_clickDomObj(pwb,"username,pass,3")
	This will recurse into the 'left' frame and try to click an object called 'results'
	iWeb_clickDomObj(pwb,"results","left")
	*/ 
		If	pWin	:=	iWeb_DomWin(pwb,frm) 
		{
			pWin.document.all.item[obj].click()
			d=1
			ObjRelease(pWin)
		}
		Return	d
	}
	iWeb_clickText(pwb,t,frm="")
	{
	/*********************************************************************
	pwb	-	browser object
	t 	-	text with in the link to check against
	frm -	frame reference; optionally, a comma delimited list of frames (name, id or index ) of all value can be used
	Example Usage
	The below will take a browser object and try to click an object with the text 'Click Here'
	iWeb_clickText(pwb,"Click Here")
	This will recurse into the 'left' frame and try to click an object with the text 'Contact Us'
	iWeb_clickText(pwb,"Contact Us","left")
	*/
		If	pWin	:=	iWeb_DomWin(pwb,frm) 
		{
			Loop,%	pWin.document.links.length
				If	InStr(pWin.document.links.item[A_Index-1].innertext,t)
				{
					pWin.document.links.item[A_Index-1].click()
					d=1
					Break
				}	
			ObjRelease(pWin)
		} ;;If	pWin
		Return	d
	}
	iWeb_clickHref(pwb,t,frm="")
	{
	/*********************************************************************
	pwb	-	browser object
	t 	-	text with in the href to check against
	frm -	frame reference; optionally, a comma delimited list of frames (name, id or index ) of all value can be used
	Example Usage
	This will click a link with the text below in the href attribute even if there were more following that entry in the href
	iWeb_clickHref(pwb,"javascript:alert('this was in a link')")
	This will recurse into the 'left' frame and try to click an object with the url for AutoHotkey's forum in an href
	iWeb_clickHref(pwb,"http://www.autohotkey.com/forum/","left")
	*/
		If	pWin	:=	iWeb_DomWin(pwb,frm) 
		{
			Loop,%	pWin.document.links.length
				If	InStr(pWin.document.links.item[ A_Index-1 ].href,t)
				{
					pWin.document.links.item[A_Index-1].click()
					d=1
					Break
				}	
			ObjRelease(pWin)
		} ;;If	pWin
		Return	d
	}
	iWeb_clickValue(pwb,t,frm="")
	{
	/*********************************************************************
	pwb	-	browser object
	t	-	text to match from visible button or other inputs
	frm -	frame reference; optionally, a comma delimited list of frames (name, id or index ) of all value can be used
	Example Usage
	The below will click an element that has a value attribute equal to 'Submit'
	iWeb_clickValue(pwb,"Submit")
	This will recurse into the 'left' frame and try to click an object with the value 'Enter'
	iWeb_clickValue(pwb,"Enter","left")
	*/
		If	pWin	:=	iWeb_DomWin(pwb,frm) 
		{
			Loop,%	pWin.document.all.length
				If	iWeb_inpt(itm:=pWin.document.all.item(A_Index-1)) ? InStr(pWin.document.all.item[A_Index-1].value,t) : 0
				{
					itm.click()
					ObjRelease(itm)
					d=1
					Break
				}	
				Else	ObjRelease(itm)
			ObjRelease(pWin)
		} ;;If	pWin
		Return	d
	}



;~ Functions used to interact with scripts embeded in a web page

;~ 	insert and execute a javascript statement into an exisiting document window
	iWeb_execScript(pwb,js,frm="")
	{
		If	(js && (pWin:=	iWeb_DomWin(pwb,frm)))
		{
			if IE_GetVersion() = 11
				pWin.eval(js)
			else
				pWin.execScript(js)
			ObjRelease(pWin)
		}
		Return
	}
;~ 	retreive a global variable value from a page
	iWeb_getVar(pwb,var,frm="")
	{
		If	(var && (pWin:=	iWeb_DomWin(pwb,frm)))
		{
			rslt:=	pWin.%var%
			ObjRelease(pWin)
		}
		Return rslt
	}
	
	;~ 	this helper function is really designed to return only 
	;~ 	useable un formated text that can be used within javascript
	iWeb_escape_text(txt)
	{
		
		StringReplace,txt,txt,',\',ALL
		StringReplace,txt,txt,"",\"",ALL
		;~ StringReplace,txt,txt,`.`.,`.,ALL
		StringReplace,txt,txt,`r,%a_space%,ALL
		StringReplace,txt,txt,`n,%a_space%,ALL
		StringReplace,txt,txt,`n`r,%a_space%,ALL
		StringReplace,txt,txt,%a_space%%a_space%,%a_space%,ALL
		return txt	
	}
;~ 	simply stripts html tags from a string
	iWeb_striphtml(HTML)
	{
;~ 		thanks lazlo http://www.autohotkey.com/forum/viewtopic.php?p=71935#71935
		Loop Parse, HTML, <>
			If (A_Index & 1) 
				noHTML .= A_LoopField
		Return noHTML
	}
	iWeb_UrlEncode( String )
	{
		OldFormat := A_FormatInteger
		SetFormat, Integer, H

		Loop, Parse, String
		{
			if A_LoopField is alnum
			{
				Out .= A_LoopField
				continue
			}
			Hex := SubStr( Asc( A_LoopField ), 3 )
			Out .= "%" . ( StrLen( Hex ) = 1 ? "0" . Hex : Hex )
		}

		SetFormat, Integer, %OldFormat%

		return Out
	}

	iWeb_uriDecode(str) { 
	   Loop 
		  If RegExMatch(str, "i)(?<=%)[\da-f]{1,2}", hex) 
			 StringReplace, str, str, `%%hex%, % Chr("0x" . hex), All 
		  Else Break 
	   Return, str 
	} 


;~ takes an html fragment and creates a DOM document from a string
iWeb_Txt2Doc(t)
{
	If	doc := ComObjCreate("{25336920-03F9-11CF-8FD0-00AA00686F13}") 
		doc.write(t),doc.close() 
	Return doc
}
;~ Sets a window and tab as active by the page title
;~ iWeb_Activate("Home") 
iWeb_Activate(sTitle) 
{ 
	DllCall("LoadLibrary", "str", "oleacc.dll") 
	HWND:=iWeb_GetWin(sTitle).HWND
	DetectHiddenWindows, On 
	WinActivate,% "ahk_id " HWND
	WinWaitActive,% "ahk_id " HWND,,5
	ControlGet, hTabBand, hWnd,, TabBandClass1, ahk_class IEFrame
	ControlGet, hTabUI  , hWnd,, DirectUIHWND1, ahk_id %hTabBand% 
	
	VarSetCapacity(CLSID, 16)
	nSize=38
	wString := sString := "{618736E0-3C3D-11CF-810C-00AA00389B71}"
	if(nSize = "")
		nSize:=DllCall("kernel32\MultiByteToWideChar", "Uint", 0, "Uint", 0, "Uint", &sString, "int", -1, "Uint", 0, "int", 0)
	VarSetCapacity(wString, nSize * 2 + 1)
	DllCall("kernel32\MultiByteToWideChar", "Uint", 0, "Uint", 0, "Uint", &sString, "int", -1, "Uint", &wString, "int", nSize + 1)
	DllCall("ole32\CLSIDFromString", "Uint",&wString , "Uint", &CLSID)
	
	If   hTabUI && DllCall("oleacc\AccessibleObjectFromWindow", "Uint", hTabUI, "Uint",-4, "Uint", &CLSID , "UintP", pacc)=0 
	{ 
		pacc := ComObject(9, pacc, 1), ObjAddRef(pacc)
		Loop, %   pacc.accChildCount 
			If   paccChild:=pacc.accChild(A_Index) 
				If   paccChild.accRole(0+0) = 0x3C 
				{ 
					paccTab:=paccChild 
					Break 
				} 
				Else   ObjRelease(paccChild) 
		ObjRelease(pacc) 
	} 
	If   pacc:=paccTab 
	{ 
		Loop, %   pacc.accChildCount
			If   paccChild:=pacc.accChild(A_Index) 
				If   paccChild.accName(0+0) = sTitle   
				{ 
					ObjRelease(pwb)
					paccChild.accDoDefaultAction(0)
					ObjRelease(paccChild) 
					Break 
				} 
				Else   ObjRelease(paccChild) 
		ObjRelease(pacc) 
	}  
	WinActivate,% sTitle
} 