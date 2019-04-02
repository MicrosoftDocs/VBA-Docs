---
title: SelectNamesDialog.Session property (Outlook)
keywords: vbaol11.chm823
f1_keywords:
- vbaol11.chm823
ms.prod: outlook
api_name:
- Outlook.SelectNamesDialog.Session
ms.assetid: 99f445e8-190b-fa26-319f-ff7783b27795
ms.date: 06/08/2017
localization_priority: Normal
---


# SelectNamesDialog.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [SelectNamesDialog](Outlook.SelectNamesDialog.md) object.


## Remarks

The  **Session** property and the **[Application.GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[SelectNamesDialog Object](Outlook.SelectNamesDialog.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]