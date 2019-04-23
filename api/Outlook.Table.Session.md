---
title: Table.Session property (Outlook)
keywords: vbaol11.chm2226
f1_keywords:
- vbaol11.chm2226
ms.prod: outlook
api_name:
- Outlook.Table.Session
ms.assetid: 8a17876d-6637-f30b-6c0f-32cfc8b77d51
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Table](Outlook.Table.md) object.


## Remarks

The  **Session** property and the **[Application.GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Table Object](Outlook.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]