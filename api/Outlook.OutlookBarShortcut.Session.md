---
title: OutlookBarShortcut.Session property (Outlook)
keywords: vbaol11.chm340
f1_keywords:
- vbaol11.chm340
ms.prod: outlook
api_name:
- Outlook.OutlookBarShortcut.Session
ms.assetid: aee32453-1650-1d28-10ae-a125fa4c4394
ms.date: 06/08/2017
localization_priority: Normal
---


# OutlookBarShortcut.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [OutlookBarShortcut](Outlook.OutlookBarShortcut.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[OutlookBarShortcut Object](Outlook.OutlookBarShortcut.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]