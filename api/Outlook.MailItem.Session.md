---
title: MailItem.Session property (Outlook)
keywords: vbaol11.chm1292
f1_keywords:
- vbaol11.chm1292
ms.prod: outlook
api_name:
- Outlook.MailItem.Session
ms.assetid: 43272ff5-ab89-f160-7995-981158f6f375
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]