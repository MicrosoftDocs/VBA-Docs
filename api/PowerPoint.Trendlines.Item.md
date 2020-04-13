---
title: Trendlines.Item method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Trendlines.Item
ms.assetid: ddda769f-ffc2-c03f-4087-755a5530f156
ms.date: 06/08/2017
localization_priority: Normal
---


# Trendlines.Item method (PowerPoint)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a '[Trendlines](PowerPoint.Trendlines.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**Variant**|The index number for the object.|

## Return value

A **[Trendline](PowerPoint.Trendline.md)** object that the collection contains.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the number of units that the trendline on the first chart in the active document extends forward and backward. The example should be run on a 2D column chart that contains a single series with a trendline.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(1).Trendlines.Item(1)

            .Forward = 5

            .Backward = .5

        End With

    End If

End With
```


## See also


[Trendlines Object](PowerPoint.Trendlines.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]