---
title: Legend.Position Property (Word)
keywords: vbawd10.chm147193989
f1_keywords:
- vbawd10.chm147193989
ms.prod: word
api_name:
- Word.Legend.Position
ms.assetid: 62d90af0-cbab-430e-3bbe-ac6058d2dfa6
ms.date: 06/08/2017
---


# Legend.Position Property (Word)

Returns or sets the position of the legend on the chart. Read/write  **[XlLegendPosition](Word.xllegendposition.md)** .


## Syntax

 _expression_ . **Position**

 _expression_ A variable that represents a **[Legend](Word.Legend.md)** object.


## Example

The following example moves the chart legend to the bottom of the chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Legend.Position = xlLegendPositionBottom 
 End If 
End With
```


## See also


#### Concepts


[Legend Object](Word.Legend.md)

