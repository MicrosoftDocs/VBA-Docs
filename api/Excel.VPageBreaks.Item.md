---
title: VPageBreaks.Item property (Excel)
keywords: vbaxl10.chm168073
f1_keywords:
- vbaxl10.chm168073
ms.prod: excel
api_name:
- Excel.VPageBreaks.Item
ms.assetid: 88e9cc81-409b-52ca-3d4e-54d3d28f186c
ms.date: 05/18/2019
localization_priority: Normal
---


# VPageBreaks.Item property (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[VPageBreaks](Excel.VPageBreaks.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number of the object.|

## Example

This example changes the location of vertical page-break one.

```vb
Worksheets(1).VPageBreaks.Item(1).Location = .Range("e5")
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]