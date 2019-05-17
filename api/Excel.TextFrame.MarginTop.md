---
title: TextFrame.MarginTop property (Excel)
keywords: vbaxl10.chm644076
f1_keywords:
- vbaxl10.chm644076
ms.prod: excel
api_name:
- Excel.TextFrame.MarginTop
ms.assetid: 5c03ceb4-e2fd-9ff7-ac5d-4fad45cd5313
ms.date: 05/17/2019
localization_priority: Normal
---


# TextFrame.MarginTop property (Excel)

Returns or sets the distance (in [points](../language/glossary/vbe-glossary.md#point)) between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text. Read/write **Single**.


## Syntax

_expression_.**MarginTop**

_expression_ A variable that represents a **[TextFrame](Excel.TextFrame.md)** object.


## Example

This example adds a rectangle to _myDocument_, adds text to the rectangle, and then sets the margins for the text frame.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 0, 0, 250, 140).TextFrame 
 .Characters.Text = "Here is some test text" 
 .MarginBottom = 0 
 .MarginLeft = 100 
 .MarginRight = 0 
 .MarginTop = 20 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]