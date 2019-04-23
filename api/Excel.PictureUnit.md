---
title: PictureUnit property (Excel Graph)
keywords: vbagr10.chm3077575
f1_keywords:
- vbagr10.chm3077575
ms.prod: excel
api_name:
- Excel.PictureUnit
ms.assetid: 28a7cd8b-2558-87a1-158f-ff9a1dca8f41
ms.date: 04/11/2019
localization_priority: Normal
---


# PictureUnit property (Excel Graph)

Returns or sets the unit for each picture on the chart if the **[PictureType](excel.picturetype.md)** property is set to **xlScale** (otherwise, this property is ignored). Read/write **Long** for all objects, except for the **[Walls](excel.walls-graph-object.md)** object, which is read/write **Variant**.

## Syntax

_expression_.**PictureUnit**

_expression_ Required. An expression that returns one of the above objects.


## Example

This example sets series one to stack pictures, and uses each picture to represent five units. The example should be run on a 2D column chart with picture data markers.

```vb
With myChart.SeriesCollection(1) 
 .PictureType = xlScale 
 .PictureUnit = 5 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]