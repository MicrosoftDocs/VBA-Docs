---
title: Axes object (Word)
ms.prod: word
api_name:
- Word.Axes
ms.assetid: 57261ca9-7fd6-ba99-19bd-5df8e940f714
ms.date: 06/08/2017
localization_priority: Normal
---


# Axes object (Word)

Represents a collection of all the  **[Axis](Word.Axis.md)** objects in the specified chart.


## Remarks

Use the  **[Axes](Word.Chart.Axes.md)** method to return the **Axes** collection.

Use  **Axes** ( _Type_ , _AxisGroup_ ), where _Type_ is the axis type and _AxisGroup_ is the axis group, to return an **Axes** collection that contains a single **Axis** object. _Type_ can be one of the following **[XlAxisType](Word.xlaxistype.md)** constants: **xlCategory**, **xlSeries**, or **xlValue**. _AxisGroup_ can be one of the following **[XlAxisGroup](Word.xlaxisgroup.md)** constants: **xlPrimary** or **xlSecondary**.


## Example

The following example displays the number of axes for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 MsgBox .Chart.Axes.Count 
 End If 
End With
```

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

## Methods

- [Item](Word.Axes.Item.md)

## Properties

- [Application](Word.Axes.Application.md)
- [Count](Word.Axes.Count.md)
- [Creator](Word.Axes.Creator.md)
- [Parent](Word.Axes.Parent.md)

## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]