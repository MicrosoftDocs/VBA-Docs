---
title: Attachments.Session property (Outlook)
keywords: vbaol11.chm172
f1_keywords:
- vbaol11.chm172
ms.prod: outlook
api_name:
- Outlook.Attachments.Session
ms.assetid: af206370-3d50-84de-187d-019126958b61
ms.date: 06/08/2017
localization_priority: Normal
---


# Attachments.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [Attachments](Outlook.Attachments.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Attachments Object](Outlook.Attachments.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]