---
title: Worksheet.StandardWidth property (Excel)
keywords: vbaxl10.chm175130
f1_keywords:
- vbaxl10.chm175130
ms.prod: excel
api_name:
- Excel.Worksheet.StandardWidth
ms.assetid: 6792ce79-0a73-fcbd-ea52-7d7aee7b9932
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.StandardWidth property (Excel)

Returns or sets the standard (default) width of all the columns on the worksheet. Read/write **Double**.


## Syntax

_expression_.**StandardWidth**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Remarks

One unit of column width is equal to the width of one character in the Normal style. For proportional fonts, the width of the character 0 (zero) is used.


## Example

This example sets the width of column one on Sheet1 to the standard width.

```vb
Worksheets("Sheet1").Columns(1).ColumnWidth = _ 
 Worksheets("Sheet1").StandardWidth
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]