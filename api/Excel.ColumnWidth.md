---
title: ColumnWidth property (Excel Graph)
keywords: vbagr10.chm65778
f1_keywords:
- vbagr10.chm65778
ms.prod: excel
api_name:
- Excel.ColumnWidth
ms.assetid: fffb3493-4b40-7a0b-f3ad-d191baebb87f
ms.date: 04/10/2019
localization_priority: Normal
---


# ColumnWidth property (Excel Graph)

Returns or sets the width of all columns in the specified range. Read/write **Variant**.

## Syntax

_expression_.**ColumnWidth**

_expression_ Required. An expression that returns a **[Range](excel.range-graph-object.md)** object.


## Remarks

One unit of column width is equal to the width of one character in the Normal style. For proportional fonts, the width of the character 0 (zero) is used.

If all columns in the range have the same width, the **ColumnWidth** property returns the width. If columns in the range have different widths, this property returns **Null**.


## Example

This example doubles the width of column A on the datasheet.

```vb
With myChart.Application.DataSheet.Columns("A") 
 .ColumnWidth = .ColumnWidth * 2 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]