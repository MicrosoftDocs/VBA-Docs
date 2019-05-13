---
title: Shape.SmartArt property (Word)
keywords: vbawd10.chm161480860
f1_keywords:
- vbawd10.chm161480860
ms.prod: word
api_name:
- Word.Shape.SmartArt
ms.assetid: d2f3fd89-288d-ac1e-18bb-00e2d043d4cd
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.SmartArt property (Word)

Returns a [SmartArt](Office.SmartArt.md) object that provides a way to work with the SmartArt associated with the specified shape. Read-only.


## Syntax

_expression_.**SmartArt**

_expression_ A variable that represents a **[Shape](Word.Shape.md)** object.


## Remarks

The  **SmartArt** property provides an entry point for interacting with a SmartArt graphic associated with the shape.


## Example

The following code example adds a SmartArt graphic to the active document.


```vb
Dim myShape As Shape 
Dim mySmartArt As SmartArt 
 
Set myShape = ActiveDocument.Shapes.AddSmartArt(Application.SmartArtLayouts(1), 100, 100, 400, 400) 
Set mySmartArt = myShape.SmartArt
```


## See also


[Shape Object](Word.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]