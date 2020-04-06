---
title: Points.Item method (Word)
keywords: vbawd10.chm10485760
f1_keywords:
- vbawd10.chm10485760
ms.prod: word
api_name:
- Word.Points.Item
ms.assetid: fae75738-6507-1b97-5179-9bc855d4c83d
ms.date: 06/08/2017
localization_priority: Normal
---


# Points.Item method (Word)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a '[Points](Word.Points.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number for the object.|

## Return value

A  **[Point](Word.Point.md)** object that the collection contains.


## Example

The following example sets the marker style for the third point in series one in embedded chart one on worksheet one. The specified series must be a 2D line, scatter, or radar series.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Points.Item(3). _ 
 MarkerStyle = xlDiamond 
 End If 
End With 

```


## See also


[Points Object](Word.Points.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]