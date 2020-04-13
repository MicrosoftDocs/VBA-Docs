---
title: DocumentItem.Session property (Outlook)
keywords: vbaol11.chm1181
f1_keywords:
- vbaol11.chm1181
ms.prod: outlook
api_name:
- Outlook.DocumentItem.Session
ms.assetid: 40c7d5d6-2efd-f946-bc2b-273209c6c896
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentItem.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [DocumentItem](Outlook.DocumentItem.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[DocumentItem Object](Outlook.DocumentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]