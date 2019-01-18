---
title: Trendline.Type property (Word)
keywords: vbawd10.chm26345580
f1_keywords:
- vbawd10.chm26345580
ms.prod: word
api_name:
- Word.Trendline.Type
ms.assetid: 1f461dcc-242e-09a5-bc63-36f1a56af82d
ms.date: 06/08/2017
localization_priority: Normal
---


# Trendline.Type property (Word)

Returns or sets the trendline type. Read/write  **[xlTrendlineType](Word.xltrendlinetype.md)**.


## Syntax

 _expression_. `Type`

 _expression_ A variable that represents a '[Trendline](Word.Trendline.md)' object.


## Example

The following example changes the trendline type for the first series of the first chart in the active document. If the series has no trendline, this example fails.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Trendlines(1).Type = xlMovingAvg 
 End If 
End With
```


## See also


[Trendline Object](Word.Trendline.md)

