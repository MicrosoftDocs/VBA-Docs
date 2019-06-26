---
title: Axes object (PowerPoint)
keywords: vbapp10.chm681000
f1_keywords:
- vbapp10.chm681000
ms.prod: powerpoint
api_name:
- PowerPoint.Axes
ms.assetid: 71f1e1fc-7086-a84e-1e05-6fa50597b49b
ms.date: 06/08/2017
localization_priority: Normal
---


# Axes object (PowerPoint)

Represents a collection of all the  **[Axis](PowerPoint.Axis.md)** objects in the specified chart.


## Remarks

Use the  **[Axes](PowerPoint.Chart.Axes.md)** method to return the **Axes** collection.

Use  **Axes** ( _Type_, _AxisGroup_ ), where _Type_ is the axis type and _AxisGroup_ is the axis group, to return an **Axes** collection that contains a single **Axis** object. _Type_ can be one of the following **[XlAxisType](PowerPoint.XlAxisType.md)** constants: **xlCategory**, **xlSeries**, or **xlValue**. _AxisGroup_ can be one of the following **[XlAxisGroup](PowerPoint.XlAxisGroup.md)** constants: **xlPrimary** or **xlSecondary**.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example displays the number of axes for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        MsgBox .Chart.Axes.Count

    End If

End With
```




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the category axis title text for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlCategory)

            .HasTitle = True

            .AxisTitle.Caption = "1994"

        End With

    End If

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]