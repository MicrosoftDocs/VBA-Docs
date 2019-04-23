---
title: OutlookBarGroup.Session property (Outlook)
keywords: vbaol11.chm323
f1_keywords:
- vbaol11.chm323
ms.prod: outlook
api_name:
- Outlook.OutlookBarGroup.Session
ms.assetid: eb75d479-7217-51b3-6426-53ff960e9c60
ms.date: 06/08/2017
localization_priority: Normal
---


# OutlookBarGroup.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [OutlookBarGroup](Outlook.OutlookBarGroup.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[OutlookBarGroup Object](Outlook.OutlookBarGroup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]