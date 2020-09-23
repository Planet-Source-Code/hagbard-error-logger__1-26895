<div align="center">

## Error Logger


</div>

### Description

This is the code I use to keep a track of errors throughout the program. It comes in useful for keeping track of the errors I haven't weeded out in the initial testing.
 
### More Info
 
erDesc - Error Description

erNum - Error Number

subname - Sub Name where error originates

Appends the error log file with the error description,number, Sub name and Date/Time


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Hagbard](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/hagbard.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__1-26.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/hagbard-error-logger__1-26895/archive/master.zip)





### Source Code

```
'Declare a Public constant ErrorLog filename and a Application Name
Public Const conErrorLogFile = "errorlog.txt"
Public const conAppname = "ApplicationName"
'The Sub for writing to the error log file
'Send Variables by value :
'erDesc = Error Description
'erNum = Error Number
'subName = name of procedure where error originated
Public Sub UpdateErrorLog(ByVal erDesc, ByVal erNum, ByVal subName As String)
On Error GoTo handleel
Dim strErrorLogname As String	'Full pathname for error log
Dim strAF As String		'Path to Application folder
  strAF = App.Path
'check for root path
  If Right(strAF, 1) = "\" Then
    strAF = strAF
    Else
    strAF = strAF & "\"
    End If
'get the full pathname from the app folder and the constant error log name
  strErrorLogname = strAF & conErrorLogFile
'open the error log file for appending - file is created if it doesn't already exist
  Open strErrorLogname For Append As #23 ' Open file for input.
  'write the line to the file
  'writes error description, error number, Sub name, Date/Time
  Write #23, erDesc, erNum, subName, Now
  'Close the file
  Close #23  ' Close file.
ExitEL:
  Exit Sub
handleel:
  MsgBox Err.Description & vbCrLf & Err.Number, , conAppName & " - Error Writing Log"
  Resume ExitEL
End Sub
'in the error handler for each sub I added the line to call the UpdateErrorLog procedure
'eg in the Form_load procedure
Private Sub Form_load()
On error goto HandleFormLoaderr1
'form_load code
'do whatever has to be done
exitloader:
  Exit Sub
HandleformLoaderr1:
  'Give the user a message
  MsgBox Err.Description & vbCrLf & Err.Number, , conAppName & " - Error"
  'Send the error info to the sub
  UpdateErrorLog Err.Description, Err.Number, "Form_Load"
  Resume exitloader
End Sub
```

