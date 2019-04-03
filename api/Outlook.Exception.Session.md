---
title: Exception.Session property (Outlook)
keywords: vbaol11.chm299
f1_keywords:
- vbaol11.chm299
ms.prod: outlook
api_name:
- Outlook.Exception.Session
ms.assetid: b8663ef0-1042-e3c4-81ca-76d4b76a3351
ms.date: 06/08/2017
localization_priority: Normal
---


# Exception.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [Exception](Outlook.Exception.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Exception Object](Outlook.Exception.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]