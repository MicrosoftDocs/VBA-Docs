---
title: NavigationFolder.Session property (Outlook)
keywords: vbaol11.chm2902
f1_keywords:
- vbaol11.chm2902
ms.prod: outlook
api_name:
- Outlook.NavigationFolder.Session
ms.assetid: f31a9538-4ebe-80f1-aa93-4d7de8e0bb7e
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationFolder.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [NavigationFolder](Outlook.NavigationFolder.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[NavigationFolder Object](Outlook.NavigationFolder.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]