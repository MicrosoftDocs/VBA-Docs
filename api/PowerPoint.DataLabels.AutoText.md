---
title: DataLabels.AutoText property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabels.AutoText
ms.assetid: 6e964058-3cfa-ba02-b324-fc1e82beb3d3
ms.date: 06/08/2017
localization_priority: Normal
---


# DataLabels.AutoText property (PowerPoint)

 **True** if all objects in the collection automatically generate appropriate text based on context. Read/write **Boolean**.


## Syntax

_expression_.**AutoText**

_expression_ A variable that represents a '[DataLabels](PowerPoint.DataLabels.md)' object.


## Remarks

Setting the value of this property sets the  **[AutoText](PowerPoint.DataLabel.AutoText.md)** property of all **[DataLabel](PowerPoint.DataLabel.md)** objects contained by the collection. This property returns **True** only when the **AutoText** property for all **DataLabel** objects contained in the collection is set to **True**; otherwise, this property returns **False**.


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


[DataLabels Object](PowerPoint.DataLabels.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]