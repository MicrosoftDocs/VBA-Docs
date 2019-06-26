---
title: Axes.Item method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Axes.Item
ms.assetid: 61657765-2c92-5fdf-c3a9-0c75ca70fe68
ms.date: 06/08/2017
localization_priority: Normal
---


# Axes.Item method (PowerPoint)

Returns a single  **[Axis](PowerPoint.Axis.md)** object from an **Axes** collection.


## Syntax

_expression_.**Item** (_Type_, _AxisGroup_)

_expression_ A variable that represents an '[Axes](PowerPoint.Axes.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**[XlAxisType](PowerPoint.XlAxisType.md)**|The axis type.|
| _AxisGroup_|Optional|**[XlAxisGroup](PowerPoint.XlAxisGroup.md)**|The axis.|

## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the title text for the category axis for the first chart in the active document.




```vb
With ActivePresentation.Slides(1).Shapes(1)

    If .HasChart Then

        With .Chart.Axes.Item(xlCategory)

            .HasTitle = True

            .AxisTitle.Caption = "1994"

        End With

    End If

End With
```


## See also


[Axes Object](PowerPoint.Axes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]