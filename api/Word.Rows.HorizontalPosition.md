---
title: Rows.HorizontalPosition property (Word)
keywords: vbawd10.chm155975695
f1_keywords:
- vbawd10.chm155975695
ms.prod: word
api_name:
- Word.Rows.HorizontalPosition
ms.assetid: 249389cb-c21f-61f2-c12a-648f70fe5357
ms.date: 06/08/2017
localization_priority: Normal
---


# Rows.HorizontalPosition property (Word)

Returns or sets the horizontal distance between the edge of the rows and the item specified by the **RelativeHorizontalPosition** property. Read/write **Single**.


## Syntax

_expression_. `HorizontalPosition`

_expression_ A variable that represents a **[Rows](Word.Rows.md)** object.


## Remarks

This property can be a number that indicates a measurement in points, or can be one of the **WdTablePosition** constants. This property doesn't have any effect if the **[WrapAroundText](Word.Rows.WrapAroundText.md)** property is **False**.


## Example

This example aligns the first table in the active document horizontally with the right margin.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 With ActiveDocument.Tables(1).Rows 
 .RelativeHorizontalPosition = _ 
 wdRelativeHorizontalPositionMargin 
 .HorizontalPosition = wdTableRight 
 End With 
End If
```


## See also


[Rows Collection Object](Word.rows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]