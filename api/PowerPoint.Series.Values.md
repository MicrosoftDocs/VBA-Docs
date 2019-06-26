---
title: Series.Values property (PowerPoint)
keywords: vbapp10.chm65700
f1_keywords:
- vbapp10.chm65700
ms.prod: powerpoint
api_name:
- PowerPoint.Series.Values
ms.assetid: ff6ceb5c-e7c3-6b75-8225-d18dd3baa2b8
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.Values property (PowerPoint)

Returns or sets a collection of all the values in the series. Read/write  **Variant**.


## Syntax

_expression_.**Values**

_expression_ A variable that represents a '[Series](PowerPoint.Series.md)' object.


## Remarks

The value of this property can be either the address of a range on the chart's worksheet or an array of constant values, but not a combination of both. See the examples for details.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the series values from a range address.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Values = "=Sheet1!B2:B5"

    End If

End With
```




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

To assign a constant value to each individual data point, you must use an array, as shown in the following example.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1).Values = _
            Array(1, 3, 5, 7, 11, 13, 17, 19)
    End If
End With
```


## See also


[Series Object](PowerPoint.Series.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]