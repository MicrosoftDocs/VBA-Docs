---
title: Axes.Item method (Word)
ms.prod: word
api_name:
- Word.Axes.Item
ms.assetid: 143898d3-cbc8-ebfc-4e25-caceeb91a8bf
ms.date: 06/08/2017
localization_priority: Normal
---


# Axes.Item method (Word)

Returns a single  **[Axis](Word.Axis.md)** object from an **Axes** collection.


## Syntax

_expression_.**Item** (_Type_, _AxisGroup_)

_expression_ A variable that represents an '[Axes](Word.Axes.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[XlAxisType](Word.xlaxistype.md)**|One of the enumeration values that specifies the axis type.|
| _AxisGroup_|Optional| **[XlAxisGroup](Word.xlaxisgroup.md)**|One of the enumeration values that specifies the axis.|

## Example

The following example sets the title text for the category axis for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes.Item(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Caption = "1994" 
 End With 
 End If 
End With
```


## See also


[Axes Object](Word.Axes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]