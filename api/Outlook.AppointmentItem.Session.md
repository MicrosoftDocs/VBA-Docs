---
title: AppointmentItem.Session property (Outlook)
keywords: vbaol11.chm840
f1_keywords:
- vbaol11.chm840
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.Session
ms.assetid: ff92a5eb-5a5a-9211-c247-42b9d993780f
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]