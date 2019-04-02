---
title: Borders.JoinBorders property (Word)
keywords: vbawd10.chm154927130
f1_keywords:
- vbawd10.chm154927130
ms.prod: word
api_name:
- Word.Borders.JoinBorders
ms.assetid: e25f3192-469e-ef65-e412-098d5cfb6173
ms.date: 06/08/2017
localization_priority: Normal
---


# Borders.JoinBorders property (Word)

 **True** if vertical borders at the edges of paragraphs and tables are removed so that the horizontal borders can connect to the page border. Read/write **Boolean**.


## Syntax

_expression_. `JoinBorders`

 _expression_ An expression that returns a '[Borders](Word.borders.md)' object.


## Example

This example adds a border around each page in the first section of the selection. The example also removes the horizontal borders at the edges of tables and paragraphs, thus connecting the horizontal borders to the page border.


```vb
Dim borderLoop As Border 
 
With Selection.Sections(1) 
 For Each borderLoop In .Borders 
 borderLoop.ArtStyle = wdArtBalloonsHotAir 
 borderLoop.ArtWidth = 15 
 Next borderLoop 
 With .Borders 
 .DistanceFromLeft = 2 
 .DistanceFromRight = 2 
 .DistanceFrom = wdBorderDistanceFromText 
 .JoinBorders = True 
 End With 
End With
```


## See also


[Borders Collection Object](Word.borders.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]