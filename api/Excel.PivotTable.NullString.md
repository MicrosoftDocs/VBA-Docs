---
title: PivotTable.NullString property (Excel)
keywords: vbaxl10.chm235114
f1_keywords:
- vbaxl10.chm235114
api_name:
- Excel.PivotTable.NullString
ms.assetid: f9d678d1-5e9f-8d3b-1f9a-73e8679ae499
ms.date: 05/09/2019
ms.localizationpriority: medium
---


# PivotTable.NullString property (Excel)

Returns or sets the string displayed in cells that contain null values when the **[DisplayNullString](Excel.PivotTable.DisplayNullString.md)** property is **True**. The default value is an empty string (""). Read/write **String**.


## Syntax

_expression_.**NullString**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Example

This example causes the PivotTable report to display "NA" in cells that contain null values.

```vb
With Worksheets(1).PivotTables("Pivot1") 
 .NullString = "NA" 
 .DisplayNullString = True 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]