---
title: Worksheet.Columns property (Excel)
keywords: vbaxl10.chm175086
f1_keywords:
- vbaxl10.chm175086
ms.prod: excel
api_name:
- Excel.Worksheet.Columns
ms.assetid: 41c18561-2a87-b975-e212-97f39fe10393
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.Columns property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents all the columns on the specified worksheet.


## Syntax

_expression_.**Columns**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Remarks

Using the **Columns** property without an object qualifier is equivalent to using **ActiveSheet.Columns**. If the active document isn't a worksheet, the **Columns** property fails.

To return a single column, use the **[Item](Excel.Range.Item.md)** property or equivalently include an index in parentheses. For example, `Columns(1)`, `Columns("A")`, `Columns.Item(1)` and `Columns.Item("A")` return the first column of the active sheet.

## Example

This example formats the font of column one (column A) on Sheet1 as bold.

```vb
Worksheets("Sheet1").Columns(1).Font.Bold = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
