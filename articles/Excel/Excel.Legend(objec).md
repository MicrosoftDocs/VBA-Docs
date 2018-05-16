---
title: Legend Object (Excel)
keywords: vbaxl10.chm621072
f1_keywords:
- vbaxl10.chm621072
ms.prod: excel
api_name:
- Excel.Legend
ms.assetid: 9be53984-bc9c-f964-9ab3-be52d3699bd9
ms.date: 06/08/2017
---


# Legend Object (Excel)

Represents the legend in a chart. Each chart can have only one legend.


## Remarks

 The **Legend** object contains one or more **[LegendEntry](Excel.LegendEntry(objec).md)** objects; each **LegendEntry** object contains a **[LegendKey](Excel.LegendKey(objec).md)** object.

The chart legend isn't visible unless the  **[HasLegend](Excel.Chart.HasLegend.md)** property is **True**. If this property is **False**, properties and methods of the **Legend** object will fail.


## Example

Use the  **[Legend](Excel.Chart.Legend.md)** property to return the **Legend** object. The following example sets the font style for the legend in embedded chart one on worksheet one to bold.


```
Worksheets(1).ChartObjects(1).Chart.Legend.Font.Bold = True
```


## Methods



|**Name**|
|:-----|
|[Clear](Excel.Legend.Clear.md)|
|[Delete](Excel.Legend.Delete.md)|
|[LegendEntries](Excel.Legend.LegendEntries.md)|
|[Select](Excel.Legend.Select.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.Legend.Application.md)|
|[Creator](Excel.Legend.Creator.md)|
|[Format](Excel.Legend.Format.md)|
|[Height](Excel.Legend.Height.md)|
|[IncludeInLayout](Excel.Legend.IncludeInLayout.md)|
|[Left](Excel.Legend.Left.md)|
|[Name](Excel.Legend.Name.md)|
|[Parent](Excel.Legend.Parent.md)|
|[Position](Excel.Legend.Position.md)|
|[Shadow](Excel.Legend.Shadow.md)|
|[Top](Excel.Legend.Top.md)|
|[Width](Excel.Legend.Width.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
