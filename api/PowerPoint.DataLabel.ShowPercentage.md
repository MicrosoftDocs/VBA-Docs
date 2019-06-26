---
title: DataLabel.ShowPercentage property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabel.ShowPercentage
ms.assetid: 00b28ebe-a674-93a1-2c6d-f8fb7d0539cf
ms.date: 06/08/2017
localization_priority: Normal
---


# DataLabel.ShowPercentage property (PowerPoint)

 **True** to display the percentage value for the data labels on a chart. **False** to hide the value. Read/write **Boolean**.


## Syntax

_expression_.**ShowPercentage**

_expression_ A variable that represents a '[DataLabel](PowerPoint.DataLabel.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables the percentage value to be shown for the data labels of the first series on the first chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).DataLabels. _
            ShowPercentage = True

    End If

End With
```


## See also


[DataLabel Object](PowerPoint.DataLabel.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]