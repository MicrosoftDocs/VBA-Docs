---
title: Shapes.Paste method (Publisher)
keywords: vbapb10.chm2162724
f1_keywords:
- vbapb10.chm2162724
ms.prod: publisher
api_name:
- Publisher.Shapes.Paste
ms.assetid: 435dd253-ae35-1dcf-ae5a-d7dfd40abf33
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.Paste method (Publisher)

Pastes the shapes or text on the Clipboard into the specified **Shapes** collection at the top of the z-order. Each pasted object becomes a member of the specified **Shapes** collection. 

If the Clipboard contains a text range, the text will be pasted into a newly created **TextFrame** shape. Returns a **[ShapeRange](Publisher.ShapeRange.md)** object that represents the pasted objects.


## Syntax

_expression_.**Paste**

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


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

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]