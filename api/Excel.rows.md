---
title: Rows property (Excel Graph)
keywords: vbagr10.chm5207942
f1_keywords:
- vbagr10.chm5207942
ms.prod: excel
ms.assetid: 045405b7-3f7c-bcf6-7757-f116ed8d7e37
ms.date: 04/12/2019
localization_priority: Normal
---


# Rows property (Excel Graph)

Returns a **Range** object that represents the rows in the specified **Range** or **[DataSheet](excel.datasheet-graph-object.md)** object. Read-only.

## Syntax

_expression_.**Rows**

_expression_ Required. An expression that returns a **[Range](excel.range-graph-object.md)** object.

## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../excel/Concepts/Workbooks-and-Worksheets/returning-an-object-from-a-collection-excel.md).

## Example

This example deletes row three on the datasheet.

```vb
myChart.Application.DataSheet.Rows(3).Delete
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]