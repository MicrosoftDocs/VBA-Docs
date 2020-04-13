---
title: NavigationModule.Session property (Outlook)
keywords: vbaol11.chm2805
f1_keywords:
- vbaol11.chm2805
ms.prod: outlook
api_name:
- Outlook.NavigationModule.Session
ms.assetid: 7fd04cbc-37c2-56e7-68b2-e7e8340cd99c
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationModule.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

 _expression_ An expression that returns a [NavigationModule](Outlook.NavigationModule.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[NavigationModule Object](Outlook.NavigationModule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]