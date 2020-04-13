---
title: Stores.Session property (Outlook)
keywords: vbaol11.chm816
f1_keywords:
- vbaol11.chm816
ms.prod: outlook
api_name:
- Outlook.Stores.Session
ms.assetid: aea9466c-4b22-10fa-7938-d12f4f193148
ms.date: 06/08/2017
localization_priority: Normal
---


# Stores.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Stores](Outlook.Stores.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Stores Object](Outlook.Stores.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]