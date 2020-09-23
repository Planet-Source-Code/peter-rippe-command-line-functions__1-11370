<div align="center">

## Command Line Functions


</div>

### Description

4 functions to handle command line options...example: yourprogy.exe -F, etc...one of them checks if a commandline option is in the command line....another one creates a commandline set with options...another one removes a specified option from the command line, and return the cmdline...(example: sometext -F, to sometext)...and the last one pulls out the text between 2 options in the commandline, and returns it...so -S blah -E would return just "blah"..can be very usefull...
 
### More Info
 
one of them returns wether the command option specified is in the command line, one returns the command line with a specified option removed, and one returns a generated command line...you specify the input, and what ever options you want.

And the last one will pull the text between 2 options and display it.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Peter Rippe](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/peter-rippe.md)
**Level**          |Beginner
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/peter-rippe-command-line-functions__1-11370/archive/master.zip)





### Source Code

```
Public Function GetCmdOpt(cmdline As String, optname As String) As Boolean
'Returns True, or False if the option (optname)
'Is inside of the commandline (cmdline).
retval2 = cmdline
retval2 = InStr(retval2, optname)
If retval2 > 0 Then
GetCmdOpt = True
Else
GetCmdOpt = False
End If
End Function
Public Function RemoveOpt(cmdline As String, optname As String) As String
'Removes a option from a commandline specified in
'cmdline...optname is the option to be removed.
retval1 = cmdstr
retval1 = Replace(retval1, optname, "")
retval1 = Trim(retval1)
RemoveOpt = retval1
End Function
Public Function AddCmdOpt(cmdline As String, optname As String) As String
'Use to add a option to a commandline...not sure
'how usefull that could be, but it might.
inputstr = Trim(inputstr)
inputstr = inputstr & " " & optname
AddCmdOpt = inputstr
End Function
Public Function GetCmdText(cmdline As String, startopt As String, endopt As String) As String
'Returns the text between 2 options specified...the start, and endoption...
'the cmdline option is the input commandline
'...If there is no option(s) specified, it wont do anything..
cmdline = LCase(cmdline)
startopt = LCase(startopt)
endopt = LCase(endopt)
If cmdline <> "" Or startopt <> "" Or endopt <> "" Then
startoptlen = InStr(cmdline, startopt)
endoptlen = InStr(cmdline, endopt)
If startoptlen > 0 Or endoptlen > 0 Then
retval1 = InStr(cmdline, endopt) - InStr(cmdline, startopt)
retval2 = Mid(cmdline, InStr(cmdline, startopt), retval1)
retval2 = Replace(retval2, startopt, "")
retval2 = Trim(retval2)
GetCmdText = retval2
End If
End If
End Function
'''''''''''''''''''''''''''''''''''''''''''''''
Example how to use each in a progy:
'Command gets the commandline from your program (myexe.exe thisiscmdmaterial)
retval = GetCmdOpt(Command, "-Test")
If retval = True Then
MsgBox "The Option WAS in the commandline"
Else
MsgBox "The Option WAS NOT in the commandline"
End If
retval = RemoveOpt(Command, "-Test")
MsgBox "The returned commadnline after the option was removed is: " & retval
retval = getcmdtext(Command, "-Start", "-End")
MsgBox "The text between the start,and end option was: " & retval
```

