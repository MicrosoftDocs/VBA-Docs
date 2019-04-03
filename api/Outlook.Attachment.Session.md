---
title: Attachment.Session property (Outlook)
keywords: vbaol11.chm2363
f1_keywords:
- vbaol11.chm2363
ms.prod: outlook
api_name:
- Outlook.Attachment.Session
ms.assetid: 0e4d45be-453a-a673-33ad-5087f5e26a98
ms.date: 06/08/2017
localization_priority: Normal
---


# Attachment.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [Attachment](Outlook.Attachment.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Attachment Object](Outlook.Attachment.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]