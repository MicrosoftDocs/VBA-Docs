---
title: NavigationFolders.Session property (Outlook)
keywords: vbaol11.chm2893
f1_keywords:
- vbaol11.chm2893
ms.prod: outlook
api_name:
- Outlook.NavigationFolders.Session
ms.assetid: 3a173fc8-3924-31f6-d0ed-967eb57089c3
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationFolders.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [NavigationFolders](Outlook.NavigationFolders.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[NavigationFolders Object](Outlook.NavigationFolders.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]