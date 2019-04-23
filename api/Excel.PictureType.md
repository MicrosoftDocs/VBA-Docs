---
title: PictureType property (Excel Graph)
keywords: vbagr10.chm3077574
f1_keywords:
- vbagr10.chm3077574
ms.prod: excel
api_name:
- Excel.PictureType
ms.assetid: 8d331b09-745e-863d-a32c-77a9f1448b85
ms.date: 04/11/2019
localization_priority: Normal
---


# PictureType property (Excel Graph)

Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3D chart. 

For the **[Point](excel.point-graph-object.md)** and **[Series](excel.series-graph-object.md)** objects, read/write **[XlChartPictureType](excel.xlchartpicturetype.md)**. Use the **[PictureUnit](Excel.PictureUnit.md)** property to determine what unit each picture represents.

For the **[LegendKey](excel.legendkey-graph-object.md)** object, read/write **Long**. 

For the **[Floor](excel.floor-graph-object.md)** and **[Walls](excel.walls-graph-object.md)** objects, read/write **Variant**.

## Syntax

_expression_.**PictureType**

_expression_ Required. An expression that returns one of the above objects.

## Example

This example sets series one to stretch pictures. The example should be run on a 2D column chart with picture data markers.

```vb
myChart.SeriesCollection(1).PictureType = xlStretch
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]