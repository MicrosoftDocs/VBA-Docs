---
title: Series.PictureType property (Word)
keywords: vbawd10.chm123732129
f1_keywords:
- vbawd10.chm123732129
ms.prod: word
api_name:
- Word.Series.PictureType
ms.assetid: 29150e44-0815-9e6e-7fcb-92f030f3cf6a
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.PictureType property (Word)

Returns or sets a value that specifies how pictures are displayed on a column or bar picture chart. Read/write  **[XlChartPictureType](Word.xlchartpicturetype.md)**.


## Syntax

_expression_.**PictureType**

_expression_ A variable that represents a '[Series](Word.Series.md)' object.


## Example

The following example sets series one of the first chart in the active document to stretch pictures. You should run the example on a 2D column chart that has picture data markers.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).PictureType = xlStretch 
 End If 
End With
```


## See also


[Series Object](Word.Series.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]