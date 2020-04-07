---
title: Chart.Axes method (Word)
ms.prod: word
api_name:
- Word.Chart.Axes
ms.assetid: 37f422b5-31f2-92ce-c04e-a837b0a3d407
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Axes method (Word)

Returns a collection of axes on the chart.


## Syntax

_expression_.**Axes** (_Type_, _AxisGroup_)

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Optional| **Variant**|The axis to return. Can be one of the following  **[XlAxisType](Word.xlaxistype.md)** constants: **xlValue**, **xlCategory**, or **xlSeriesAxis** (**xlSeriesAxis** is valid only for 3D charts).|
| _AxisGroup_|Optional| **[XlAxisGroup](Word.xlaxisgroup.md)**|One of the enumeration values that specifies the axis group. The default is **xlPrimary**.
> [!NOTE] 
> 3D charts have only one axis group.

|

## Return value

An [Axes](Word.Axes.md) object that contains the selected axes from the chart.


## Example

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


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]