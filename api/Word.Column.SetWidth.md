---
title: Column.SetWidth method (Word)
keywords: vbawd10.chm156172489
f1_keywords:
- vbawd10.chm156172489
ms.prod: word
api_name:
- Word.Column.SetWidth
ms.assetid: fd42d86d-53a4-c05d-81c3-add15cf05766
ms.date: 06/08/2017
localization_priority: Normal
---


# Column.SetWidth method (Word)

Sets the width of a column in a table.


## Syntax

_expression_. `SetWidth`( `_ColumnWidth_` , `_RulerStyle_` )

_expression_ Required. A variable that represents a '[Column](Word.Column.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ColumnWidth_|Required| **Single**|The width of the specified column or columns, in points.|
| _RulerStyle_|Required| **WdRulerStyle**|Controls the way Word adjusts cell widths.|

## Remarks

The  **[WdRulerStyle](Word.WdRulerStyle.md)** behavior described above applies to left-aligned tables. The **WdRulerStyle** behavior for center- and right-aligned tables can be unexpected; in these cases, the **SetWidth** method should be used with care.


## See also


[Column Object](Word.Column.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]