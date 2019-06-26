---
title: DataLabel.ShowValue property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabel.ShowValue
ms.assetid: 2d4ca0a0-9b2c-7477-214b-322283e2c082
ms.date: 06/08/2017
localization_priority: Normal
---


# DataLabel.ShowValue property (PowerPoint)

 **True** to display a specified chart's data label values. **False** to hide the values. Read/write **Boolean**.


## Syntax

_expression_.**ShowValue**

_expression_ A variable that represents a '[DataLabel](PowerPoint.DataLabel.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables the value to be shown for the data labels of the first series in the first chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then
        .Chart.SeriesCollection(1).DataLabels. _
            ShowValue = True
    End If

End With
```


## See also


[DataLabel Object](PowerPoint.DataLabel.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]