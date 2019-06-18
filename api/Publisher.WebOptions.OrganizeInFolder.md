---
title: WebOptions.OrganizeInFolder property (Publisher)
keywords: vbapb10.chm8257542
f1_keywords:
- vbapb10.chm8257542
ms.prod: publisher
api_name:
- Publisher.WebOptions.OrganizeInFolder
ms.assetid: f09ac701-d8d8-a58f-965c-bd5e4b69820c
ms.date: 06/18/2019
localization_priority: Normal
---


# WebOptions.OrganizeInFolder property (Publisher)

Returns or sets a **Boolean** value that specifies whether a web publication is saved in a flat structure or hierarchical structure. If **False**, all files in the web publication are saved in a flat structure within the root folder. If **True**, the files are saved in a hierarchical structure within the root folder. The default value is **True**. Read/write.


## Syntax

_expression_.**OrganizeInFolder**

_expression_ A variable that represents a **[WebOptions](Publisher.WebOptions.md)** object.


## Return value

Boolean


## Example

The following example specifies that all files in the web publication should be saved in a flat structure within the root folder.

```vb
Dim theWO As WebOptions 
 
Set theWO = Application.WebOptions 
 
With theWO 
 .OrganizeInFolder = False 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]