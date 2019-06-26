---
title: DataLabel.AutoText property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabel.AutoText
ms.assetid: f7e154ad-4f5f-0a3d-3fe5-c83994705cfb
ms.date: 06/08/2017
localization_priority: Normal
---


# DataLabel.AutoText property (PowerPoint)

 **True** if the object automatically generates appropriate text based on context. Read/write **Boolean**.


## Syntax

_expression_.**AutoText**

_expression_ A variable that represents a '[DataLabel](PowerPoint.DataLabel.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the data labels for series one of the first chart in the active document to automatically generate appropriate text.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1). _
            DataLabels.AutoText = True
    End If
End With
```


## See also


[DataLabel Object](PowerPoint.DataLabel.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]