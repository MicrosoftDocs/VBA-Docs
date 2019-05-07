---
title: PivotTable.DisplayNullString property (Excel)
keywords: vbaxl10.chm235105
f1_keywords:
- vbaxl10.chm235105
ms.prod: excel
api_name:
- Excel.PivotTable.DisplayNullString
ms.assetid: ad2ce480-9fc9-d069-5526-4f819e236967
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.DisplayNullString property (Excel)

**True** if the PivotTable report displays a custom string in cells that contain null values. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**DisplayNullString**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

Use the **[NullString](Excel.PivotTable.NullString.md)** property to set the custom null string.


## Example

This example causes the PivotTable report to display "NA" in cells that contain null values.

```vb
With Worksheets(1).PivotTables("Pivot1") 
 .NullString = "NA" 
 .DisplayNullString = True 
End With
```

<br/>

This example causes the PivotTable report to display 0 (zero) in cells that contain null values.

```vb
Worksheets(1).PivotTables("Pivot1").DisplayNullString = False
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]