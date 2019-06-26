---
title: DataLabel.ShowCategoryName property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabel.ShowCategoryName
ms.assetid: 7eeb3ab4-d0e3-3682-0ea4-a75fae60b800
ms.date: 06/08/2017
localization_priority: Normal
---


# DataLabel.ShowCategoryName property (PowerPoint)

 **True** to display the category name for the data labels on a chart. **False** to hide the category name. Read/write **Boolean**.


## Syntax

_expression_.**ShowCategoryName**

_expression_ A variable that represents a '[DataLabel](PowerPoint.DataLabel.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example shows the category name for the data labels of the first series on the first chart.




```vb
With ActiveDocument.InlineShapes(1) 
    If .HasChart Then 
        .Chart.SeriesCollection(1).DataLabels. _ 
            ShowCategoryName = True 
    End If 
End With
```


## See also


[DataLabel Object](PowerPoint.DataLabel.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]