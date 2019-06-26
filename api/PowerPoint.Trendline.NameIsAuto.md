---
title: Trendline.NameIsAuto property (PowerPoint)
keywords: vbapp10.chm65724
f1_keywords:
- vbapp10.chm65724
ms.prod: powerpoint
api_name:
- PowerPoint.Trendline.NameIsAuto
ms.assetid: 7fe8b6ef-b5d9-5a97-64b2-561552654684
ms.date: 06/08/2017
localization_priority: Normal
---


# Trendline.NameIsAuto property (PowerPoint)

 **True** if Microsoft Word automatically determines the name of the trendline. Read/write **Boolean**.


## Syntax

_expression_. `NameIsAuto`

_expression_ A variable that represents a '[Trendline](PowerPoint.Trendline.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets Microsoft Word to automatically determine the name for trendline one of the first chart in the active document. You should run the example on a 2D column chart that contains a single series that has a trendline.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1) _
            .Trendlines(1).NameIsAuto = True
    End If
End With
```


## See also


[Trendline Object](PowerPoint.Trendline.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]