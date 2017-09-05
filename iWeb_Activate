iWeb_Activate(sTitle) 
{ 
	DllCall("LoadLibrary", "str", "oleacc.dll") 
	HWND:=IEGet(sTitle).HWND
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
