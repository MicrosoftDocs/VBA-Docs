---
title: TableView.Session property (Outlook)
keywords: vbaol11.chm2501
f1_keywords:
- vbaol11.chm2501
ms.prod: outlook
api_name:
- Outlook.TableView.Session
ms.assetid: 6443565e-2a7a-5466-a68e-9baf13e316c5
ms.date: 06/08/2017
localization_priority: Normal
---


# TableView.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [TableView](Outlook.TableView.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[TableView Object](Outlook.TableView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]