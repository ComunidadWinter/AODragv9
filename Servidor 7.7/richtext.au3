#include <GUIConstants.au3>
#include <GuiConstantsEx.au3>
#include <GUIedit.au3>
#include <Misc.au3>



WinActivate("Server", "")
WinWaitActive("Server", "")


$texto = ControlGetText("", "", "[CLASS:RichTextWndClass; INSTANCE:4]")
MsgBox(0, "", $texto)
Exit




;~ FUNCIONES
;~ Gettok como el del mIRC
Func _GetTok($_TokList,$_TokNum,$_ChrMatch)
    Local $_ChrCheck = _Iif(IsNumber($_ChrMatch),Chr($_ChrMatch),$_ChrMatch)
    Dim $a = 0, $_List = StringSplit($_TokList,$_ChrCheck)
    Return $_List[$_TokNum]
EndFunc
