---
title: Series.PictureType property (Excel)
keywords: vbaxl10.chm578101
f1_keywords:
- vbaxl10.chm578101
ms.prod: excel
api_name:
- Excel.Series.PictureType
ms.assetid: 098dac46-ec2d-ea2d-71e9-1094a5f0b23a
ms.date: 05/11/2019
localization_priority: Normal
---


# Series.PictureType property (Excel)

Returns or sets an **[XlChartPictureType](Excel.XlChartPictureType.md)** value that represents the way pictures are displayed on a column or bar picture chart.


## Syntax

_expression_.**PictureType**

_expression_ A variable that represents a **[Series](Excel.Series(object).md)** object.


## Example

This example sets series one on Chart1 to stretch pictures. The example should be run on a 2D column chart with picture data markers.

```vb
Charts("Chart1").SeriesCollection(1).PictureType = xlStretch
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]