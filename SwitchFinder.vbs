set shell = CreateObject("WScript.Shell")
set oFSO = CreateObject("Scripting.FileSystemObject")
Dim exe, exepath, outpath

'**************EDIT HERE*************************
exe="test.exe" 
exepath="C:\PathToExe"
outpath="C:\PathToOutputFile"
testargs=Array("a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z","help","quiet","silent")
ecart=10 'ecart permit en octets entre 2 réponses pour qu'elles soient considérées comme identiques
'***********************************************
reffile= outpath & "ref.txt"
argfile= outpath & "args.txt"
tmpfile= outpath & "tmp.txt"

shell.run("cmd /c " & exepath & exe & " -refrefref > " & outpath & "ref.txt") 'genere la reponse d'argument invalide de reference
WScript.Sleep 1000
shell.run "taskkill /f /im " & exe, , True
set oFile=oFSO.GetFile(reffile)
tailleref=oFile.Size ' taille de reference



For Each arg in testargs
	shell.run("cmd /c " & exepath & exe & " -" & arg & " > " & outpath & "tmp.txt") 
	WScript.Sleep 1000
	shell.run "taskkill /f /im " & exe, , True
	set oFile=oFSO.GetFile(tmpfile)
	taillearg=oFile.Size
	If Abs(taillearg-tailleref)> ecart Then
		shell.run("cmd /c echo -" & arg & " >> " & argfile) 
	End If
Next