---
title: Points.Item method (Excel)
keywords: vbaxl10.chm574075
f1_keywords:
- vbaxl10.chm574075
ms.prod: excel
api_name:
- Excel.Points.Item
ms.assetid: 1e588b64-3676-63ab-5136-eec028a82a4e
ms.date: 05/09/2019
localization_priority: Normal
---


# Points.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[Points](Excel.Points(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number for the object.|

## Return value

A **[Point](Excel.Point(object).md)** object contained by the collection.


## Example

This example sets the marker style for the third point in series one in embedded chart one on worksheet one. The specified series must be a 2D line, scatter, or radar series.

```vb
Worksheets(1).ChartObjects(1).Chart. _ 
 SeriesCollection(1).Points.Item(3).MarkerStyle = xlDiamond
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]