---
title: Rectangle object (Word)
keywords: vbawd10.chm3571
f1_keywords:
- vbawd10.chm3571
ms.prod: word
api_name:
- Word.Rectangle
ms.assetid: 90ad4f48-2051-38f9-9b2e-a14bd38478be
ms.date: 06/08/2017
localization_priority: Normal
---


# Rectangle object (Word)

Represents a portion of text or a graphic in a page. Use the **Rectangle** object and related methods and properties for programmatically defining page layout in a document.


## Remarks

Use the **Item** method to return a specific **Rectangle** object. The following example accesses the first rectangle in the first page of the active document.


```vb
Dim objRectangle As Rectangle 
 
Set objRectangle = ActiveDocument.ActiveWindow _ 
 .Panes(1).Pages(1).Rectangles.Item(1)
```

Use the **RectangleType** property to determine the type of rectangle. The following example creates a **ShapeRange** object if the specified rectangle is a shape.




```vb
Dim objRectangle As Rectangle 
Dim objShape As ShapeRange 
 
Set objRectangle = ActiveDocument.ActiveWindow _ 
 .Panes(1).Pages(1).Rectangles.Item(1) 
 
If objRectangle.RectangleType = wdShapeRectangle Then 
 Set objShape = objRectangle.Range.ShapeRange 
End If
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]