---
title: DataTable Object (Excel)
keywords: vbaxl10.chm625072
f1_keywords:
- vbaxl10.chm625072
ms.prod: excel
api_name:
- Excel.DataTable
ms.assetid: aca0850b-2e72-cde9-b751-633876e1df99
ms.date: 06/08/2017
---


# DataTable Object (Excel)

Represents a chart data table.


## Example

Use the  **[DataTable](Excel.Chart.DataTable.md)** property to return a **DataTable** object. The following example adds a data table with an outline border to embedded chart one.


```
With Worksheets(1).ChartObjects(1).Chart 
 .HasDataTable = True 
 .DataTable.HasBorderOutline = True 
End With
```


## Methods



|**Name**|
|:-----|
|[Delete](Excel.DataTable.Delete.md)|
|[Select](Excel.DataTable.Select.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.DataTable.Application.md)|
|[Border](Excel.DataTable.Border.md)|
|[Creator](Excel.DataTable.Creator.md)|
|[Font](Excel.DataTable.Font.md)|
|[Format](Excel.DataTable.Format.md)|
|[HasBorderHorizontal](Excel.DataTable.HasBorderHorizontal.md)|
|[HasBorderOutline](Excel.DataTable.HasBorderOutline.md)|
|[HasBorderVertical](Excel.DataTable.HasBorderVertical.md)|
|[Parent](Excel.DataTable.Parent.md)|
|[ShowLegendKey](Excel.DataTable.ShowLegendKey.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
