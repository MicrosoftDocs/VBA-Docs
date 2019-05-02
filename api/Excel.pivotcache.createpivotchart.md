---
title: PivotCache.CreatePivotChart method (Excel)
keywords: vbaxl10.chm227110
f1_keywords:
- vbaxl10.chm227110
ms.prod: excel
ms.assetid: 5aeb9a16-2cf8-3525-12b0-0b6e3d3ddf1a
ms.date: 05/03/2019
localization_priority: Normal
---


# PivotCache.CreatePivotChart method (Excel)

Creates a standalone PivotChart from a **PivotCache** object. Returns a **[Shape](Excel.Shape.md)** object.


## Syntax

_expression_.**CreatePivotChart** (_ChartDestination_, _XlChartType_, _Left_, _Top_, _Width_, _Height_)

_expression_ A variable that represents a **[PivotCache](Excel.PivotCache.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ChartDestination_|Required|**Variant**|The Destination worksheet.|
| _XlChartType_|Optional|**Variant**|The type of chart.|
| _Left_|Optional|**Variant**|The distance, in [points](../language/glossary/vbe-glossary.md#point), from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).|
| _Top_|Optional|**Variant**|The distance, in points, from the top edge of the topmost shape in the shape range to the top edge of the worksheet.|
| _Width_|Optional|**Variant**|The width, in points, of the object.|
| _Height_|Optional|**Variant**|The height, in points, of the object.|

## Return value

**Shape** object


## Remarks

If the **PivotCache** object that the method is called from has no attached PivotTable:

- A workbook-level PivotTable is created from the existing PivotCache.
    
- A standalone PivotChart is created with a reference to the newly created PivotTable.
    
If the PivotCache already has an associated PivotTable:

- The PivotCache is cloned.
    
- A new workbook-level PivotTable is created based on the cloned PivotCache.
    
- A standalone PivotChart is created with a reference to the new workbook-level PivotTable.
    

## Example

The following code creates a decoupled PivotChart from a PivotCache object.

```vb
Workbooks("Book1").Connections.Add _
     "cubes4 Adventure Works DW 2008 Special Char Adventure Works", "", Array( _
     "OLEDB;Provider=MSOLAP.4;Integrated Security=SSPI;Persist Security Info=True;Data Source=<server name here >;Initial Catalog=Adventure Works DW 2008" _
     , " Special Char"), Array("Adventure Works"), 1
   ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
     ActiveWorkbook.Connections( _
     "cubes4 Adventure Works DW 2008 Special Char Adventure Works"), Version:= _
     xlPivotTableVersion14).CreatePivotChart(ChartDestination:="Sheet1").Select

   ActiveChart.ChartType = xlColumnClustered
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]