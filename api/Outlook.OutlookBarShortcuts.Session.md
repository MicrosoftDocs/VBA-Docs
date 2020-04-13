---
title: OutlookBarShortcuts.Session property (Outlook)
keywords: vbaol11.chm331
f1_keywords:
- vbaol11.chm331
ms.prod: outlook
api_name:
- Outlook.OutlookBarShortcuts.Session
ms.assetid: 538cc6e5-2772-23bb-6ed4-658ed8607660
ms.date: 06/08/2017
localization_priority: Normal
---


# OutlookBarShortcuts.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [OutlookBarShortcuts](Outlook.OutlookBarShortcuts.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[OutlookBarShortcuts Object](Outlook.OutlookBarShortcuts.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]