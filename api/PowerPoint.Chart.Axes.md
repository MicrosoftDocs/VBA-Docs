---
title: Chart.Axes method (PowerPoint)
keywords: vbapp10.chm684016
f1_keywords:
- vbapp10.chm684016
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.Axes
ms.assetid: 6f740a9e-2baa-5a84-ea51-6a39452e227e
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Axes method (PowerPoint)

Returns a collection of axes on the chart.


## Syntax

_expression_.**Axes** (_Type_, _AxisGroup_)

_expression_ A variable that represents a **[Chart](PowerPoint.Chart.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Optional|**Variant**|The axis to return. Can be one of the following  **[XlAxisType](PowerPoint.XlAxisType.md)** constants: **xlValue**, **xlCategory**, or **xlSeriesAxis** (**xlSeriesAxis** is valid only for 3D charts).|
| _AxisGroup_|Optional|**[XlAxisGroup](PowerPoint.XlAxisGroup.md)**|One of the enumeration values that specifies the axis group. The default is **xlPrimary**.
> [!NOTE] 
> 3D charts have only one axis group.

|

## Return value

An [Axes](PowerPoint.Axes.md) object that contains the selected axes from the chart.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example adds an axis label to the category axis for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1) 
    If .HasChart Then 
        With .Chart.Axes(xlCategory) 
            .HasTitle = True 
            .AxisTitle.Text = "July Sales" 
        End With 
    End If 
End With
```

The following example turns off major gridlines in the category axis for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1) 
    If .HasChart Then 
        .Chart.Axes(xlCategory). _ 
            HasMajorGridlines = False 
    End If 
End With
```

The following example turns off all gridlines for all axes in the first chart of the active document.




```vb
With ActiveDocument.InlineShapes(1) 
    If .HasChart Then 
        For Each a In .Chart.Axes 
            a.HasMajorGridlines = False 
            a.HasMinorGridlines = False 
        Next 
    End If 
End With
```


## See also


[Chart Object](PowerPoint.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]