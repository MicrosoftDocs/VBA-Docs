---
title: ViewFields.Session property (Outlook)
keywords: vbaol11.chm2548
f1_keywords:
- vbaol11.chm2548
ms.prod: outlook
api_name:
- Outlook.ViewFields.Session
ms.assetid: 480ac826-b966-9204-8850-214b53a1b0da
ms.date: 06/08/2017
localization_priority: Normal
---


# ViewFields.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [ViewFields](Outlook.ViewFields.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[ViewFields Object](Outlook.ViewFields.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]