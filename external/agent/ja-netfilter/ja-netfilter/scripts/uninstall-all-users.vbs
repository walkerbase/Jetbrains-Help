If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName, """" & WScript.ScriptFullName & """ /elevate", "", "runas", 10
  WScript.Quit
End If

MsgBox "It may take a few seconds to execute this script." & vbCrLf & vbCrLf & "Click 'OK' button and wait for the prompt of 'Done.' to pop up!"

Sub RemoveEnv(env)
	On Error Resume Next

	Dim sEnvKey, sEnvVal, aJBProducts
	aJBProducts = Array("appcode", "aqua", "clion", "datagrip", "dataspell", "devecostudio", "gateway", "goland", "idea", "jetbrains_client", "jetbrainsclient", "phpstorm", "pycharm", "rider", "rubymine", "rustrover", "studio", "webide", "webstorm")

	For Each sPrd in aJBProducts
    	sEnvKey = UCase(sPrd) & "_VM_OPTIONS"
    	sEnvVal = oShell.ExpandEnvironmentStrings("%" & sEnvKey & "%")
    	If sEnvVal <> ("%" & sEnvKey & "%") Then
        	env.Remove(sEnvKey)
    	End If
		Next
End Sub

Set oShell = CreateObject("WScript.Shell")

RemoveEnv oShell.Environment("USER")
RemoveEnv oShell.Environment("SYSTEM")

MsgBox "Done."
