---
title: Document.Shapes property (Word)
keywords: vbawd10.chm158007358
f1_keywords:
- vbawd10.chm158007358
ms.prod: word
api_name:
- Word.Document.Shapes
ms.assetid: 638ab04b-2e82-afe9-3817-740f464542cc
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Shapes property (Word)

Returns a  **[Shapes](Word.shapes.md)** collection that represents all the **Shape** objects in the specified document. Read-only.


## Syntax

_expression_.**Shapes**

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

This collection can contain drawings, shapes, pictures, OLE objects, ActiveX controls, text objects, and callouts. For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).

The **Shapes** property, when applied to a document, returns all the **Shape** objects in the main story of the document, excluding the headers and footers.


## Example

This example creates a new document, adds a rectangle to it that's 100 points wide and 50 points high, and sets the upper-left corner of the rectangle to be 5 points from the left edge and 25 points from the upper-left corner of the page.


```vb
Set myDoc = Documents.Add 
myDoc.Shapes.AddShape msoShapeRectangle, 5, 25, 100, 50
```

This example sets the fill texture for all the shapes in the active document.




```vb
For Each s in ActiveDocument.Shapes 
 s.Fill.PresetTextured msoTextureOak 
Next s
```

This example adds a shadow to the first shape in the active document.




```vb
Set myShape = ActiveDocument.Shapes(1) 
myShape.Shadow.Type = msoShadow6
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]