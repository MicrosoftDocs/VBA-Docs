---
title: Points.Item method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Points.Item
ms.assetid: d3a6b3cf-3fbb-1e0f-b9cf-0b707839de67
ms.date: 06/08/2017
localization_priority: Normal
---


# Points.Item method (PowerPoint)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a '[Points](PowerPoint.Points.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|The index number for the object.|

## Return value

A **[Point](PowerPoint.Point.md)** object that the collection contains.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the marker style for the third point in series one in embedded chart one on worksheet one. The specified series must be a 2D line, scatter, or radar series.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1).Points.Item(3). _
            MarkerStyle = xlDiamond
    End If
End With
```


## See also


[Points Object](PowerPoint.Points.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]