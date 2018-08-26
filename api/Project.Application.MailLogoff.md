---
title: Application.MailLogoff Method (Project)
ms.prod: project-server
api_name:
- Project.Application.MailLogoff
ms.assetid: e8634331-404c-6e01-4ce9-2dac8dcf364c
ms.date: 06/08/2017
---


# Application.MailLogoff Method (Project)

Closes an established MAPI mail session.


## Syntax

 _expression_. `MailLogoff`

 _expression_ A variable that represents an [Application](./Project.Application.md) object.


### Return Value

Nothing


## Example

The following example checks for an existing mail session and logs off it. If not logged on, the following example logs on, downloads any new mail, and then logs off.


```vb
Sub LogoffFromMail() 
 
 If Not IsNull(MailSession) Then 
<<<<<<< HEAD
 MsgBox "Logging off mail session: " &; MailSession 
=======
 MsgBox "Logging off mail session: " & MailSession 
>>>>>>> master
 Application.MailLogoff 
 Else 
 MsgBox "Logging on to mail session now." 
 Application.MailLogon DownloadNewMail:=True 
<<<<<<< HEAD
 MsgBox "Logging off mail session: " &; MailSession 
=======
 MsgBox "Logging off mail session: " & MailSession 
>>>>>>> master
 Application.MailLogoff 
 End If 
 
End Sub
```


