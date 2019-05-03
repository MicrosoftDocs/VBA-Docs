---
title: PivotField.LayoutBlankLine property (Excel)
keywords: vbaxl10.chm240119
f1_keywords:
- vbaxl10.chm240119
ms.prod: excel
api_name:
- Excel.PivotField.LayoutBlankLine
ms.assetid: 8770c3df-0a24-0511-9e8f-44515a6772df
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.LayoutBlankLine property (Excel)

**True** if a blank row is inserted after the specified row field in a PivotTable report. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**LayoutBlankLine**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

You can set this property for any PivotTable field; however, the blank row appears only if the specified field is a row field other than the innermost (lowest-level) row field. 

For non-OLAP data sources, the value of this property doesn't change when the field is rearranged or added to the PivotTable report.

You cannot enter data in the blank row in the PivotTable report.


## Example

This example adds a blank line after the state field in the first PivotTable report on the active worksheet.

```vb
With ActiveSheet.PivotTables("PivotTable1") _ 
 .PivotFields("state") 
 .LayoutBlankLine = True 
End With
```


**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]