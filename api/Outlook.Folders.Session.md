---
title: Folders.Session property (Outlook)
keywords: vbaol11.chm41
f1_keywords:
- vbaol11.chm41
ms.prod: outlook
api_name:
- Outlook.Folders.Session
ms.assetid: 1f8d8e11-d4d9-6769-37af-5c97e1413023
ms.date: 06/08/2017
localization_priority: Normal
---


# Folders.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Folders](Outlook.Folders.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Folders Object](Outlook.Folders.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]