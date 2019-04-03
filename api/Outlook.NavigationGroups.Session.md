---
title: NavigationGroups.Session property (Outlook)
keywords: vbaol11.chm2854
f1_keywords:
- vbaol11.chm2854
ms.prod: outlook
api_name:
- Outlook.NavigationGroups.Session
ms.assetid: b742bee6-7067-8168-ebd9-2823da65dd0f
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationGroups.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [NavigationGroups](Outlook.NavigationGroups.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[NavigationGroups Object](Outlook.NavigationGroups.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]