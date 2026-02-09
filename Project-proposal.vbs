X=MsgBox("Error - File cannot be opened.",0+16,"File System Error")
X=MsgBox("The file you are trying to open is corrupted or contains invalid data.",0+64,"Cannot Proceed")
X=MsgBox("Do you want to attempt automatic repair?",0+16,"Repair File")
X=MsgBox("Repair failed. File is severely damaged.",0+64,"Repair Failed")
X=MsgBox("Sorry but the file really can't be opened po :((",0+64,"xntyyy")
X=MsgBox("Just kidding!",0+64,"xntyyy")
X=MsgBox("Happy Valentine's, Aly!",0+64,"xntyyy")


Dim shell
Set shell = CreateObject("WScript.Shell")

shell.Run "codes\index.html", 1, False

Set shell = Nothing