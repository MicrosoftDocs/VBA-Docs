---
title: RecentFiles.Item property (Excel)
keywords: vbaxl10.chm172075
f1_keywords:
- vbaxl10.chm172075
ms.prod: excel
api_name:
- Excel.RecentFiles.Item
ms.assetid: f153bdeb-6c13-2ea8-506a-2b762b211c67
ms.date: 05/11/2019
localization_priority: Normal
---


# RecentFiles.Item property (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[RecentFiles](Excel.RecentFiles.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number of the object.|

## Example

This example opens file two in the list of recently used files.

```vb
Application.RecentFiles.Item(2).Open
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]