---
title: Borders.HasVertical property (Word)
keywords: vbawd10.chm154927132
f1_keywords:
- vbawd10.chm154927132
ms.prod: word
api_name:
- Word.Borders.HasVertical
ms.assetid: dc99eb20-3bc3-2ee9-b6d6-f9a9c1b4e880
ms.date: 06/08/2017
localization_priority: Normal
---


# Borders.HasVertical property (Word)

 **True** if a vertical border can be applied to the specified object. Read-only **Boolean**.


## Syntax

 _expression_. `HasVertical`

 _expression_ A variable that represents a '[Borders](Word.borders.md)' object.


## Remarks

Vertical borders can be applied to ranges that contain cells in two or more columns of a table.


## Example

If the selection supports vertical borders, this example applies a single vertical border.


```vb
If Selection.Borders.HasVertical = True Then 
 Selection.Borders(wdBorderVertical).LineStyle = _ 
 wdLineStyleSingle 
End If
```


## See also


[Borders Collection Object](Word.borders.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]