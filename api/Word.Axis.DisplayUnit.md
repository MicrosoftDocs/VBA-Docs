---
title: Axis.DisplayUnit property (Word)
keywords: vbawd10.chm113049671
f1_keywords:
- vbawd10.chm113049671
ms.prod: word
api_name:
- Word.Axis.DisplayUnit
ms.assetid: b3f8bbbb-d532-679a-fbb1-01260554425e
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.DisplayUnit property (Word)

Returns or sets the unit label for the value axis. Read/write  **[xlDisplayUnit](Word.xldisplayunit.md)** , **xlCustom** , or **xlNone**.


## Syntax

 _expression_. `DisplayUnit`

 _expression_ A variable that represents an '[Axis](Word.Axis.md)' object.


## Remarks

Using unit labels when charting large values makes your tick-mark labels easier to read. For example, if you label your value axis in units of hundreds, thousands, or millions, you can use smaller numeric values at the tick marks on the axis.


## Example

The following example sets the units displayed on the value axis of the first chart in the active document to hundreds.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlValue) 
 .DisplayUnit = xlHundreds 
 .HasTitle = True 
 .AxisTitle.Caption = "Rebate Amounts" 
 End With 
 End If 
End With
```


## See also


[Axis Object](Word.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]