---
title: Application.MailSystem method (Project)
ms.prod: project-server
api_name:
- Project.Application.MailSystem
ms.assetid: 4ee9011c-f5f5-d0aa-0cd6-aa90130af4af
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.MailSystem method (Project)

Returns the type of email system installed on the host machine.


## Syntax

_expression_. `MailSystem`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Return value

[PjMailSystem](Project.PjMailSystem.md)


## Remarks

Can return one of the [PjMailSystem](Project.PjMailSystem.md) constants.


## Example

The following example sends the project file if the host machine is using MAPI.


```vb
Sub SendMAPI() 
 
 If Application.MailSystem = pjMAPI Then 
 MailSend To:="Jean Selva", Subject:="Sample Subject" 
 End If 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]