---
title: PostItem.Session property (Outlook)
keywords: vbaol11.chm1511
f1_keywords:
- vbaol11.chm1511
ms.prod: outlook
api_name:
- Outlook.PostItem.Session
ms.assetid: 53dc4396-598e-197b-cea1-135e44686b91
ms.date: 06/08/2017
localization_priority: Normal
---


# PostItem.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [PostItem](Outlook.PostItem.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[PostItem Object](Outlook.PostItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]