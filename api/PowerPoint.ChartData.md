---
title: ChartData object (PowerPoint)
keywords: vbapp10.chm689000
f1_keywords:
- vbapp10.chm689000
ms.prod: powerpoint
api_name:
- PowerPoint.ChartData
ms.assetid: b7bedf0e-5f11-001d-a97c-e8d07939bc8b
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartData object (PowerPoint)

Represents access to the linked or embedded data associated with a chart.


## Remarks

Use the  **[ChartData](PowerPoint.Chart.ChartData.md)** property to return the **ChartData** object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example uses the  **[Activate](PowerPoint.ChartData.Activate.md)** method to display the data associated with the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1).Chart.ChartData

    .Activate

End With
```


## Methods



|Name|
|:-----|
|[Activate](PowerPoint.ChartData.Activate.md)|
|[ActivateChartDataWindow](PowerPoint.chartdata.activatechartdatawindow.md)|
|[BreakLink](PowerPoint.ChartData.BreakLink.md)|

## Properties



|Name|
|:-----|
|[IsLinked](PowerPoint.ChartData.IsLinked.md)|
|[Workbook](PowerPoint.ChartData.Workbook.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]