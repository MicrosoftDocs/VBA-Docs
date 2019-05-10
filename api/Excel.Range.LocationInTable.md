---
title: Range.LocationInTable property (Excel)
keywords: vbaxl10.chm144156
f1_keywords:
- vbaxl10.chm144156
ms.prod: excel
api_name:
- Excel.Range.LocationInTable
ms.assetid: 7a86a0fe-cd46-331e-595b-6be168091d0c
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.LocationInTable property (Excel)

Returns a constant that describes the part of the **[PivotTable](Excel.PivotTable.md)** report that contains the upper-left corner of the specified range. Can be one of the following **[XlLocationInTable](Excel.XlLocationInTable.md)** constants. Read-only **Long**.


## Syntax

_expression_.**LocationInTable**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Example

This example displays a message box that describes the location of the active cell within the PivotTable report.

```vb
Worksheets("Sheet1").Activate 
Select Case ActiveCell.LocationInTable 
Case Is = xlRowHeader 
 MsgBox "Active cell is part of a row header" 
Case Is = xlColumnHeader 
 MsgBox "Active cell is part of a column header" 
Case Is = xlPageHeader 
 MsgBox "Active cell is part of a page header" 
Case Is = xlDataHeader 
 MsgBox "Active cell is part of a data header" 
Case Is = xlRowItem 
 MsgBox "Active cell is part of a row item" 
Case Is = xlColumnItem 
 MsgBox "Active cell is part of a column item" 
Case Is = xlPageItem 
 MsgBox "Active cell is part of a page item" 
Case Is = xlDataItem 
 MsgBox "Active cell is part of a data item" 
Case Is = xlTableBody 
 MsgBox "Active cell is part of the table body" 
End Select
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]