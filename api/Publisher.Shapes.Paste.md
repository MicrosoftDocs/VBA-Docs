---
title: Shapes.Paste Method (Publisher)
keywords: vbapb10.chm2162724
f1_keywords:
- vbapb10.chm2162724
ms.prod: publisher
api_name:
- Publisher.Shapes.Paste
ms.assetid: 435dd253-ae35-1dcf-ae5a-d7dfd40abf33
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.Paste Method (Publisher)

Pastes the shapes or text on the Clipboard into the specified  **[Shapes](Publisher.Shapes.md)** collection, at the top of the z-order. Each pasted object becomes a member of the specified **Shapes** collection. If the Clipboard contains a text range, the text will be pasted into a newly created **TextFrame** shape. Returns a **[ShapeRange](Publisher.ShapeRange.md)** object that represents the pasted objects.


## Syntax

 _expression_. **Paste**

 _expression_ A variable that represents a  **Shapes** object.


## Return value

ShapeRange


## Example

This example copies shape one on page one in the active publication to the Clipboard and then pastes it into page two.


```vb
With ActiveDocument 
 .Pages(1).Shapes(1).Copy 
 .Pages(2).Shapes.Paste 
End With
```


