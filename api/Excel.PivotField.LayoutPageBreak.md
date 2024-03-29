---
title: PivotField.LayoutPageBreak property (Excel)
keywords: vbaxl10.chm240121
f1_keywords:
- vbaxl10.chm240121
api_name:
- Excel.PivotField.LayoutPageBreak
ms.assetid: 3b513f5c-353b-ecb9-16c4-6e1d4bd0848a
ms.date: 05/04/2019
ms.localizationpriority: medium
---


# PivotField.LayoutPageBreak property (Excel)

**True** if a page break is inserted after each field. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**LayoutPageBreak**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

Although you can set this property for any PivotTable field, the print option appears only if the specified field is a row field other than the innermost (lowest-level) row field. 

For non-OLAP data sources, the value of this property doesn't change when the field is rearranged or when it is added to or removed from the PivotTable report.


## Example

This example adds a page break after the state field in the first PivotTable report on the active worksheet.

```vb
With ActiveSheet.PivotTables("PivotTable1") _ 
 .PivotFields("state") 
 .LayoutPageBreak = True 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]