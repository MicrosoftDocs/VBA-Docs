---
title: Worksheet.Rows property (Excel)
keywords: vbaxl10.chm175122
f1_keywords:
- vbaxl10.chm175122
ms.prod: excel
api_name:
- Excel.Worksheet.Rows
ms.assetid: 5d07304e-a3c9-2a75-b2ba-4a7b16ce6516
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.Rows property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents all the rows on the specified worksheet. 


## Syntax

_expression_.**Rows**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Remarks

Using the **Rows** property without an object qualifier is equivalent to using **ActiveSheet.Rows**. If the active document isn't a worksheet, the **Rows** property fails.

To return a single row, use the **[Item](Excel.Range.Item.md)** property or equivalently include an index in parentheses. For example, both `Rows(1)` and `Rows.Item(1)` return the first row of the active sheet.

## Example

This example deletes row three on Sheet1.

```vb
Worksheets("Sheet1").Rows(3).Delete
```

<br/>

This example deletes all rows on worksheet one where the value of cell one in the row is the same as the value of cell one in the previous row.

```vb
For Each rw In Worksheets(1).Rows 
   this = rw.Cells(1, 1).Value 
   If this = last Then rw.Delete 
   last = this 
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
