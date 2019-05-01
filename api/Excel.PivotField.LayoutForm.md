---
title: PivotField.LayoutForm property (Excel)
keywords: vbaxl10.chm240122
f1_keywords:
- vbaxl10.chm240122
ms.prod: excel
api_name:
- Excel.PivotField.LayoutForm
ms.assetid: 5e0fee89-111f-0bd4-e880-72cc0925c364
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotField.LayoutForm property (Excel)

Returns or sets the way the specified PivotTable items appear—in table format or in outline format. Read/write  **[XlLayoutFormType](Excel.XlLayoutFormType.md)**.


## Syntax

_expression_.**LayoutForm**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks



| **xlLayoutFormType** can be one of these **xlLayoutFormType** constants.|
| **xlTabular**. Default.|
| **xlOutline**. The **[LayoutSubtotalLocation](Excel.PivotField.LayoutSubtotalLocation.md)** property specifies where the subtotal appears in the PivotTable report.|

You can set this property for any PivotTable field; however, the formatting appears only if the specified field is a row field other than the innermost (lowest-level) row field. For non-OLAP data sources, the value of this property doesn't change when the field is rearranged or when it is added to or removed from the PivotTable report.


## Example

This example displays the state field in the first PivotTable report on the active worksheet in outline format, and it displays the subtotals at the top of the field.


```vb
With ActiveSheet.PivotTables("PivotTable1") _ 
 .PivotFields("state") 
 .LayoutForm = xlOutline 
 .LayoutSubtotalLocation = xlTop 
End With
```


## See also


[PivotField Object](Excel.PivotField.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]