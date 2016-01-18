Function lockFile(Path) 

	Set objFSO				= CreateObject("Scripting.FileSystemObject") 
	tFile					= objFSO.GetTempName()
	tPath					= Left(Path,InStrRev(Path,"\")) & tFile
 
	On Error Resume Next 
		objFSO.MoveFile Path,tPath 
	On Error GoTo 0
	
	WScript.Sleep 500
	If objFSO.FileExists(tPath) Then lockFile = tFile

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
theFilePath					= GETENV("FILE_PATH")
Set theTrash 				= objShell.NameSpace(10)
theTimeout					= 1000 * 30
theClock					= 0

Do
	WScript.Sleep 500	
	theClock = theClock + 500		
	theTempFileName			= lockFile(theFilePath)	
	If (theTempFileName <> "") Or (theClock > theTimeout) Then Exit Do
Loop

Set theFolder				= objShell.NameSpace(theFolderPath)
Set theFile					= theFolder.ParseName(theTempFileName)

theTrash.MoveHere theFile, 0

Do
	WScript.Sleep 500		
	theClock = theClock + 500		
	If (theFolder.ParseName(theTempFileName) is Nothing) Or (theClock > theTimeout) Then Exit Do
Loop


	Set objWshShell 		= Nothing