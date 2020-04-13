---
title: Rows.VerticalPosition property (Word)
keywords: vbawd10.chm155975697
f1_keywords:
- vbawd10.chm155975697
ms.prod: word
api_name:
- Word.Rows.VerticalPosition
ms.assetid: 5f08e24f-6b0c-441d-c067-41b83b4ec1c3
ms.date: 06/08/2017
localization_priority: Normal
---


# Rows.VerticalPosition property (Word)

Returns or sets the vertical distance between the edge of the rows and the item specified by the **RelativeVerticalPosition** property. Read/write **Single**.


## Syntax

_expression_. `VerticalPosition`

_expression_ Required. A variable that represents a **[Rows](Word.Rows.md)** object.


## Remarks

This property can be a number that indicates a measurement in points, or can be any valid  **[WdTablePosition](Word.WdTablePosition.md)** constant.


## Example

This example vertically aligns the first table in the active document with the top of the page.


```vb
Set myTable = ActiveDocument.Tables(1).Rows 
With myTable 
 .RelativeVerticalPosition = wdRelativeVerticalPositionPage 
 .VerticalPosition = wdTableTop 
End With
```


## See also


[Rows Collection Object](Word.rows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]