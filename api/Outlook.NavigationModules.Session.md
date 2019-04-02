---
title: NavigationModules.Session property (Outlook)
keywords: vbaol11.chm2797
f1_keywords:
- vbaol11.chm2797
ms.prod: outlook
api_name:
- Outlook.NavigationModules.Session
ms.assetid: ce7f293c-cce6-5471-fd41-3387c2f0195e
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationModules.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

 _expression_ An expression that returns a [NavigationModules](Outlook.NavigationModules.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[NavigationModules Object](Outlook.NavigationModules.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]