---
title: Range.UseStandardWidth property (Excel)
keywords: vbaxl10.chm144214
f1_keywords:
- vbaxl10.chm144214
ms.prod: excel
api_name:
- Excel.Range.UseStandardWidth
ms.assetid: 970e3d68-3147-a52f-b831-ae7780c735e0
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.UseStandardWidth property (Excel)

**True** if the column width of the **Range** object equals the standard width of the sheet. Returns **null** if the range contains more than one column and the columns aren't all the same width. Read/write **Variant**.


## Syntax

_expression_.**UseStandardWidth**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Example

This example sets the width of column A on Sheet1 to the standard width.

```vb
Worksheets("Sheet1").Columns("A").UseStandardWidth = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]