---
title: Application.MailSession method (Project)
ms.prod: project-server
api_name:
- Project.Application.MailSession
ms.assetid: 00f67414-eb0d-6b2a-d557-26812aaee04c
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.MailSession method (Project)

Returns the MAPI mail session number as a hexadecimal string if there is an active session, or returns  **Null** if there is no session.


## Syntax

_expression_. `MailSession`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Return value

 **String**


## Example

The following example gets the MAPI mail session number.


```vb
Sub Mail_Session() 
 
 Dim Return_MAPI As String 
 Return_MAPI = MailSession() 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]