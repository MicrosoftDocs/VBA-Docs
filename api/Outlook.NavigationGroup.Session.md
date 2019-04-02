---
title: NavigationGroup.Session property (Outlook)
keywords: vbaol11.chm2884
f1_keywords:
- vbaol11.chm2884
ms.prod: outlook
api_name:
- Outlook.NavigationGroup.Session
ms.assetid: 8be45a52-1a91-2b89-567d-051e1a99178c
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationGroup.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [NavigationGroup](Outlook.NavigationGroup.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[NavigationGroup Object](Outlook.NavigationGroup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]