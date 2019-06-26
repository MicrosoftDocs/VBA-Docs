---
title: Series.XValues property (PowerPoint)
keywords: vbapp10.chm66647
f1_keywords:
- vbapp10.chm66647
ms.prod: powerpoint
api_name:
- PowerPoint.Series.XValues
ms.assetid: e1e83dc0-ed73-c29b-942a-575511ce94e1
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.XValues property (PowerPoint)

Returns or sets an array of x values for a chart series. Read/write  **Variant**.


## Syntax

_expression_.**XValues**

_expression_ A variable that represents a '[Series](PowerPoint.Series.md)' object.


## Remarks

You can set the  **XValues** property to a range on a worksheet or to an array of values, but not to a combination of both.

For PivotChart reports, this property is read-only.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the values from a range address.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).XValues = "=Sheet1!B1:B5"

    End If

End With
```




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

To assign a constant value to each individual data point, you must use an array, as shown in the following example.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1).XValues = _
            Array(5.0, 6.3, 12.6, 28, 50)
    End If
End With
```


## See also


[Series Object](PowerPoint.Series.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]