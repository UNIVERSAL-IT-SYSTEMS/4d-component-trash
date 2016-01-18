Function lockFolder(Path) 

	Set objFSO				= CreateObject("Scripting.FileSystemObject") 
	tPath					= Left(Path,InStrRev(Path,"\")) & objFSO.GetTempName() 
 
	On Error Resume Next 
		objFSO.MoveFolder Path,tPath 
	On Error GoTo 0
	
	WScript.Sleep 500
	If objFSO.FolderExists(tPath) Then lockFolder = tPath

	Set objFSO				= Nothing
	
End Function


Function GETENV(variableName)
	
	Set objWshShell 		= WScript.CreateObject("WScript.Shell")
	Set WshSysEnv			= objWshShell.Environment("PROCESS")
	GETENV					= WshSysEnv(variableName)
	Set objWshShell 		= Nothing

end Function

Set objShell 				= CreateObject("Shell.Application")
theFolderPath				= GETENV("FOLDER_PATH")
Set theTrash 				= objShell.NameSpace(10)
theTimeout					= 1000 * 30
theClock					= 0
Do
	WScript.Sleep 500
	theClock = theClock + 500	
	theTempFolderPath	= lockFolder(theFolderPath)	
	If (theTempFolderPath <> "") Or (theClock > theTimeout) Then Exit Do
Loop	

Set theFolder			= objShell.NameSpace(theTempFolderPath)

theTrash.MoveHere theFolder.Self, 0
		
Do
	WScript.Sleep 500
	theClock = theClock + 500			
	If (objShell.NameSpace(theTempFolderPath) is Nothing) Or (theClock > theTimeout) Then Exit Do
Loop

	Set objWshShell 		= Nothing
