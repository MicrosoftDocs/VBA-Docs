---
title: CubeField.LayoutForm property (Excel)
keywords: vbaxl10.chm668087
f1_keywords:
- vbaxl10.chm668087
ms.prod: excel
api_name:
- Excel.CubeField.LayoutForm
ms.assetid: a9077651-214f-6926-89fc-c29a1ff35682
ms.date: 04/23/2019
localization_priority: Normal
---


# CubeField.LayoutForm property (Excel)

Returns or sets the way the specified PivotTable items appear—in table format or in outline format. Read/write **[XlLayoutFormType](Excel.XlLayoutFormType.md)**.


## Syntax

_expression_.**LayoutForm**

_expression_ A variable that represents a **[CubeField](Excel.CubeField.md)** object.


## Remarks

You can set this property for any PivotTable field; however, the formatting appears only if the specified field is a row field other than the innermost (lowest-level) row field. 

For non-OLAP data sources, the value of this property doesn't change when the field is rearranged or when it is added to or removed from the PivotTable report.


## Example

This example displays the state field in the first PivotTable report on the active worksheet in outline format, and it displays the subtotals at the top of the field.

```vb
With ActiveSheet.PivotTables("PivotTable1") _ 
 .PivotFields("state") 
 .LayoutForm = xlOutline 
 .LayoutSubtotalLocation = xlTop 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]