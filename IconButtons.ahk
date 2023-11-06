;///////////////////////////////////////////////////////////////////////////////////////////
; This script is designed to showcase creating image buttons for a GUI in AHK
; It works with multiple image types. I know it can at least use .PNG, .JPG, .JPEG, .BMP, .GIF (Will not animate)
; 
; Special thanks to Fanatic Guru who built the original function, please refer to the board topic at the link below for more information on it
;	https://www.autohotkey.com/boards/viewtopic.php?p=11390#p11390
;///////////////////////////////////////////////////////////////////////////////////////////

;Script Environment Settings
;///////////////////////////////////////////////////////////////////////////////////////////
#Persistent
#singleInstance, force
;///////////////////////////////////////////////////////////////////////////////////////////

;Prompt user about loading custom icons (Not Shell32.dll or ImageRes.dll icons)
;///////////////////////////////////////////////////////////////////////////////////////////
MsgBox, 262148, Load Custom Icons, Would you like to load Custom Icons as well?
	ifMsgBox Yes
	{
		CustomIconsOn = 1
		FileSelectFolder, CustomIconFolder, ,3, Choose where your Custom Icons are stored
	}
;///////////////////////////////////////////////////////////////////////////////////////////

;Begin loading icons and Desiging the Gui. This step takes some time because we're loading a minimum of 1,326 image buttons
;///////////////////////////////////////////////////////////////////////////////////////////
SplashTextOn, 150, 40, Loading, Loading Icons`nPlease wait..
Gui, -Theme
Gui, Tab
if (CustomIconsOn = 1)
	Gui, Add, Tab3, xm ym w1365 h985 NoTab, Shell32.dll GUI Theme On|Shell32.dll GUI Theme Off|Imageres.dll GUI Theme On|Imageres.dll GUI Theme Off|Custom Icons GUI Theme On|Custom Icons GUI Theme Off
else
	Gui, Add, Tab3, xm ym w1365 h985 NoTab, Shell32.dll GUI Theme On|Shell32.dll GUI Theme Off|Imageres.dll GUI Theme On|Imageres.dll GUI Theme Off

;-------------------------------------------------------------------------------------------
;Load Shell32 Icons with the Gui Theme ON

ypos = 30
xpos = 10
IconCount := 0
Gui, Tab, 1
Gui, +Theme
Loop, 329
{
	IconCount += 1
	Gui, Add, Text, xm+%xpos% ym+%ypos%, %IconCount%.
	xpos += 25
	Gui, Add, Button, xm+%xpos% ym+%ypos% h45 w45 hwndIcon%A_Index%
	GuiButtonIcon(Icon%A_Index%, "shell32.dll", IconCount, "h40 w40")
	xpos += 50
	if (xpos > 1350)
	{
		xpos = 10
		ypos +=50 
	}
}

;-------------------------------------------------------------------------------------------
;Load Shell32 Icons with the Gui Theme OFF

ypos = 30
xpos = 10
IconCount := 0
Gui, -Theme
Gui, Tab, 2
Loop, 329
{
	IconCount += 1
	Gui, Add, Text, xm+%xpos% ym+%ypos%, %IconCount%.
	xpos += 25
	Gui, Add, Button, xm+%xpos% ym+%ypos% h45 w45 hwndIcon%A_Index%
	GuiButtonIcon(Icon%A_Index%, "shell32.dll", IconCount, "h40 w40")
	xpos += 50
	if (xpos > 1350)
	{
		xpos = 10
		ypos +=50
	}
}

;-------------------------------------------------------------------------------------------
;Load ImageRes Icons with the Gui Theme ON

ypos = 30
xpos = 10
IconCount := 0
Gui, +Theme
Gui, Tab, 3
Loop, 334
{
	IconCount += 1
	Gui, Add, Text, xm+%xpos% ym+%ypos%, %IconCount%.
	xpos += 25
	Gui, Add, Button, xm+%xpos% ym+%ypos% h45 w45 hwndIcon%A_Index%
	GuiButtonIcon(Icon%A_Index%, "imageres.dll", IconCount, "h40 w40")
	xpos += 50
	if (xpos > 1350)
	{
		xpos = 10
		ypos +=50
	}
}

;-------------------------------------------------------------------------------------------
;Load ImageRes Icons with the Gui Theme OFF

ypos = 30
xpos = 10
IconCount := 0
Gui,-Theme
Gui, Tab, 4
Loop, 334
{
	IconCount += 1
	Gui, Add, Text, xm+%xpos% ym+%ypos%, %IconCount%.
	xpos += 25
	Gui, Add, Button, xm+%xpos% ym+%ypos% h45 w45 hwndIcon%A_Index%
	GuiButtonIcon(Icon%A_Index%, "imageres.dll", IconCount, "h40 w40")
	xpos += 50
	if (xpos > 1350)
	{
		xpos = 10
		ypos +=50
	}
}

;-------------------------------------------------------------------------------------------
;Load Custom Icons with the Gui Theme ON

if (CustomIconsOn = 1)
{
	ypos = 30
	xpos = 10
	IconCount := 0
	Gui, +Theme
	Gui, Tab, 5
	Loop, Files, %CustomIconFolder%\*.*, F ; For simplicity, this omits folders so that only files are shown in the ListView.
	{
		if (IconCount >= 342)
			Break
		if A_LoopFileExt in jpg,png,gif,jpeg,bmp,jfif
		{
			IconCount += 1
			Gui, Add, Text, xm+%xpos% ym+%ypos%, %IconCount%.
			xpos += 25
			CustomIcons := CustomIconFolder "\" A_LoopFileName
			Gui, Add, Button, xm+%xpos% ym+%ypos% h45 w45 hwndIcon%A_Index%
			GuiButtonIcon(Icon%A_Index%, CustomIcons, ,"h40 w40")
			xpos += 50
			if (xpos > 1350)
			{
				xpos = 10
				ypos +=50 
			}
		}
	}
	
;-------------------------------------------------------------------------------------------
;Load Custom Icons with the Gui Theme OFF

	ypos = 30
	xpos = 10
	IconCount := 0
	Gui, -Theme
	Gui, Tab, 6
	Loop, Files, %CustomIconFolder%\*.*, F ; For simplicity, this omits folders so that only files are shown in the ListView.
	{
		if (IconCount >= 342)
		{
			MsgBox, 262144, Maximum Reached, Custom Icon Limit Reached. The system can fit no more than 342 Icons on each tab, you may need to split your icons into multiple folders.
			Break
		}
		if A_LoopFileExt in jpg,png,gif,jpeg,bmp,jfif
		{
			IconCount += 1
			Gui, Add, Text, xm+%xpos% ym+%ypos%, %IconCount%.
			xpos += 25
			CustomIcons := CustomIconFolder "\" A_LoopFileName
			Gui, Add, Button, xm+%xpos% ym+%ypos% h45 w45 hwndIcon%A_Index%
			GuiButtonIcon(Icon%A_Index%, CustomIcons, ,"h40 w40")
			xpos += 50
			if (xpos > 1350)
			{
				xpos = 10
				ypos +=50 
			}
		}
	}
}
Gui, Show
SplashTextOff
Return

GuiEscape:
GuiClose:
Exitapp

^F3::
Reload
Return
;///////////////////////////////////////////////////////////////////////////////////////////

;Function for attaching the image to the button
;///////////////////////////////////////////////////////////////////////////////////////////
GuiButtonIcon(Handle, File, Index := 1, Options := "")
{
	RegExMatch(Options, "i)w\K\d+", W), (W="") ? W := 16 :
	RegExMatch(Options, "i)h\K\d+", H), (H="") ? H := 16 :
	RegExMatch(Options, "i)s\K\d+", S), S ? W := H := S :
	RegExMatch(Options, "i)l\K\d+", L), (L="") ? L := 0 :
	RegExMatch(Options, "i)t\K\d+", T), (T="") ? T := 0 :
	RegExMatch(Options, "i)r\K\d+", R), (R="") ? R := 0 :
	RegExMatch(Options, "i)b\K\d+", B), (B="") ? B := 0 :
	RegExMatch(Options, "i)a\K\d+", A), (A="") ? A := 4 :
	Psz := A_PtrSize = "" ? 4 : A_PtrSize, DW := "UInt", Ptr := A_PtrSize = "" ? DW : "Ptr"
	VarSetCapacity( button_il, 20 + Psz, 0 )
	NumPut( normal_il := DllCall( "ImageList_Create", DW, W, DW, H, DW, 0x21, DW, 1, DW, 1 ), button_il, 0, Ptr )	; Width & Height
	NumPut( L, button_il, 0 + Psz, DW )		; Left Margin
	NumPut( T, button_il, 4 + Psz, DW )		; Top Margin
	NumPut( R, button_il, 8 + Psz, DW )		; Right Margin
	NumPut( B, button_il, 12 + Psz, DW )	; Bottom Margin	
	NumPut( A, button_il, 16 + Psz, DW )	; Alignment
	SendMessage, BCM_SETIMAGELIST := 5634, 0, &button_il,, AHK_ID %Handle%
	return IL_Add( normal_il, File, Index )
}
;///////////////////////////////////////////////////////////////////////////////////////////

;How to use in your script
;///////////////////////////////////////////////////////////////////////////////////////////
;Step 1: Create your GUI
;Step 2: Create your Button
;Step 3: Call the GuiButtonIcon() function
;-------------------------------------------------------------------------------------------

;To use a DLL Icon
	;Gui, New
	;Gui, Add, Button, xm ym h45 w45 hwndButton1 ; I recommend NOT setting any text onto the button, if you want text I suggest using a tooltip on hover function
	;GuiButtonIcon(Button1, "shell32.dll", 1, "h40 w40") ; I recommend setting the Height/Width to be about 5 pixels smaller than your button to create a buffer and avoid clipping
	
;To use a Custom image
	;
	;Gui, New
	;Gui, Add, Button, xm ym h45 w45 hwndButton1 ; I recommend NOT setting any text onto the button, if you want text I suggest using a tooltip on hover function
	;GuiButtonIcon(Button1, "C:\image.jpg", , "h40 w40") ; I recommend setting the Height/Width to be about 5 pixels smaller than your button to create a buffer and avoid clipping			
;///////////////////////////////////////////////////////////////////////////////////////////
