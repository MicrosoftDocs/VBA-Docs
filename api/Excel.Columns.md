---
title: Columns property (Excel Graph)
keywords: vbagr10.chm65777
f1_keywords:
- vbagr10.chm65777
ms.prod: excel
api_name:
- Excel.Columns
ms.assetid: 7c5bd414-aa86-49e6-c853-0fa0c56d11a7
ms.date: 04/10/2019
localization_priority: Normal
---


# Columns property (Excel Graph)

Returns a **Range** object that represents the columns in the specified range or all the columns on the datasheet. Read-only **Range** object.

## Syntax

_expression_.**Columns**

_expression_ Required. An expression that returns a **[Range](excel.range-graph-object.md)** object.

## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../excel/Concepts/Workbooks-and-Worksheets/returning-an-object-from-a-collection-excel.md).

## Example

This example clears column A of the datasheet.

```vb
myChart.Application.DataSheet.Columns(2).ClearContents
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]