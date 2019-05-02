---
title: PageSetup.PrintTitleColumns property (Excel)
keywords: vbaxl10.chm473097
f1_keywords:
- vbaxl10.chm473097
ms.prod: excel
api_name:
- Excel.PageSetup.PrintTitleColumns
ms.assetid: 860cf212-0fbb-f3ec-c9ce-a0df57b39b7f
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.PrintTitleColumns property (Excel)

Returns or sets the columns that contain the cells to be repeated on the left side of each page, as a **String** in A1-style notation in the language of the macro. Read/write **String**.


## Syntax

_expression_.**PrintTitleColumns**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Remarks

If you specify only part of a column or columns, Microsoft Excel expands the range to full columns.

Set this property to **False** or to the empty string ("") to turn off title columns.

This property applies only to worksheet pages.


## Example

This example defines row three as the title row, and it defines columns one through three as the title columns.

```vb
Worksheets("Sheet1").Activate 
ActiveSheet.PageSetup.PrintTitleRows = ActiveSheet.Rows(3).Address 
ActiveSheet.PageSetup.PrintTitleColumns = _ 
 ActiveSheet.Columns("A:C").Address
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]