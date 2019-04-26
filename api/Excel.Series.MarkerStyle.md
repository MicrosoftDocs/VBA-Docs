---
title: Series.MarkerStyle property (Excel)
keywords: vbaxl10.chm578098
f1_keywords:
- vbaxl10.chm578098
ms.prod: excel
api_name:
- Excel.Series.MarkerStyle
ms.assetid: fec57799-b01b-a8f8-2c26-1e7b11dd9777
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.MarkerStyle property (Excel)

Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write  **[XlMarkerStyle](Excel.XlMarkerStyle.md)**.


## Syntax

_expression_.**MarkerStyle**

_expression_ A variable that represents a [Series](Excel.Series-graph-object.md) object.


## Remarks





| **xlMarkerStyle** can be one of these **xlMarkerStyle** constants.|
| **xlMarkerStyleAutomatic**. Automatic markers|
| **xlMarkerStyleCircle**. Circular markers|
| **xlMarkerStyleDash**. Long bar markers|
| **xlMarkerStyleDiamond**. Diamond-shaped markers|
| **xlMarkerStyleDot**. Short bar markers|
| **xlMarkerStyleNone**. No markers|
| **xlMarkerStylePicture**. Picture markers|
| **xlMarkerStylePlus**. Square markers with a plus sign|
| **xlMarkerStyleSquare**. Square markers|
| **xlMarkerStyleStar**. Square markers with an asterisk|
| **xlMarkerStyleTriangle**. Triangular markers|
| **xlMarkerStyleX**. Square markers with an X|

## Example

This example sets the marker style for series one in Chart1. The example should be run on a 2D line chart.


```vb
Charts("Chart1").SeriesCollection(1) _ 
 .MarkerStyle = xlMarkerStyleCircle 

```


## See also


[Series Object](Excel.Series(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
