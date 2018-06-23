#include <ScreenCapture.au3>
#include <Constants.au3>
#include <Misc.au3>

$vAppId = "nntraintool-autosave"
$vGuiDelay = 500 ; Sleep time after WinSetState and Send commands in miliseconds
$vGuiWait = 5 ; WinWaitActive command wait time in seconds

If _Singleton($vAppId, 1) = 0 Then
 Exit 9
EndIf

If $CmdLine[0] <> 1 Then
 MsgBox(16, $vAppId & ": Error", "Executable parameters not 1 (" & $CmdLine[0] & ")")
 Exit 1
EndIf

$oDictionary = ObjCreate("Scripting.Dictionary")
If @error Then
 MsgBox(16, $vAppId & ": Error", 'ObjCreate("Scripting.Dictionary") failed')
Else
 $oDictionary.Add("nntraintool.png", "[REGEXPTITLE:nntraintool]")
 $oDictionary.Add("plotperform.fig", "[REGEXPTITLE:plotperform]")
 $oDictionary.Add("plottrainstate.fig", "[REGEXPTITLE:plottrainstate]")
 $oDictionary.Add("ploterrhist.fig", "[REGEXPTITLE:ploterrhist]")
 $oDictionary.Add("plotregression.fig", "[REGEXPTITLE:plotregression]")
EndIf

$vOutputLocation = $CmdLine[1]

For $vKey In $oDictionary
 $hWindow = WinGetHandle($oDictionary.Item($vKey))
 If Not $hWindow Then
  MsgBox(16, $vAppId & ": Error", 'WinGetHandle("' & $oDictionary.Item($vKey) & '") failed')
  Exit 1
 EndIf

 $vActivate = WinActivate($hWindow)
 If Not $vActivate Then
  MsgBox(16, $vAppId & ": Error", 'WinActivate("' & $oDictionary.Item($vKey) & '") failed')
  Exit 1
 EndIf

 $vWaitActive = WinWaitActive($hWindow, "", $vGuiWait)
 If Not $vWaitActive Then
  MsgBox(16, $vAppId & ": Error", 'WinWaitActive("' & $oDictionary.Item($vKey) & '", "", ' & $vGuiWait & ') failed')
  Exit 1
 EndIf

 If $vKey = "nntraintool.png" Then
  _ScreenCapture_CaptureWnd($vOutputLocation & "nntraintool.png", $hWindow)
  If @error Then
   MsgBox(16, $vAppId & ": Error", '_ScreenCapture_CaptureWnd("' & $vOutputLocation & "nntraintool.png" & '", "' & $oDictionary.Item($vKey) & '") failed')
   Exit 1
  EndIf
 Else
  Sleep($vGuiDelay)

  $vSetState = WinSetState($hWindow, "", @SW_MAXIMIZE)
  If Not $vSetState Then
   MsgBox(16, $vAppId & ": Error", 'WinSetState("' & $oDictionary.Item($vKey) & '", "", @SW_MAXIMIZE) failed')
   Exit 1
  EndIf

  Sleep($vGuiDelay)
  Send("{ALT}")
  Sleep($vGuiDelay)
  Send("f")
  Sleep($vGuiDelay)
  Send("a")
  Sleep($vGuiDelay)

  $hSaveAs = WinWaitActive("Save As", "", $vGuiWait)
  If Not $hSaveAs Then
   MsgBox(16, $vAppId & ": Error", 'WinWaitActive("Save As", "", ' & $vGuiWait & ') failed')
   Exit 1
  EndIf

  $vSetText = ControlSetText($hSaveAs, "", "[CLASS:Edit; INSTANCE:1]", $vOutputLocation & $vKey)
  If Not $vSetText Then
   MsgBox(16, $vAppId & ": Error", 'ControlSetText("Save As", "", "[CLASS:Edit; INSTANCE:1]", "' & $vOutputLocation & $vKey & '") failed')
   Exit 1
  EndIf

  $couldSave = ControlClick($hSaveAs, "", "[CLASS:Button; INSTANCE:2]")
  If Not $couldSave Then
   MsgBox(16, $vAppId & ": Error", 'ControlClick("Save As", "", "[CLASS:Button; INSTANCE:2]") failed')
   Exit 1
  EndIf

  $vSetState = WinSetState($hWindow, "", @SW_MINIMIZE)
  If Not $vSetState Then
   MsgBox(16, $vAppId & ": Error", 'WinSetState("' & $oDictionary.Item($vKey) & '", "", @SW_MINIMIZE) failed')
   Exit 1
  EndIf

  Sleep($vGuiDelay)
 EndIf
Next

$vActivate = WinActivate("[REGEXPTITLE:MATLAB]")
If Not $vActivate Then
 MsgBox(16, $vAppId & ": Error", 'WinActivate("[REGEXPTITLE:MATLAB]") failed')
 Exit 1
EndIf

$vWaitActive = WinWaitActive($vActivate, "", $vGuiWait)
If Not $vWaitActive Then
 MsgBox(16, $vAppId & ": Error", 'WinWaitActive("[REGEXPTITLE:MATLAB]", "", ' & $vGuiWait & ') failed')
 Exit 1
EndIf

Send("{F5}")
