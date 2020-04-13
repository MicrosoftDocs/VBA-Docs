---
title: Recipient.Session property (Outlook)
keywords: vbaol11.chm2342
f1_keywords:
- vbaol11.chm2342
ms.prod: outlook
api_name:
- Outlook.Recipient.Session
ms.assetid: 0719e438-c9b0-ecca-1aa0-f25c9b21fe69
ms.date: 06/08/2017
localization_priority: Normal
---


# Recipient.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Recipient](Outlook.Recipient.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Recipient Object](Outlook.Recipient.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]