---
title: InlineShapes Object (Publisher)
keywords: vbapb10.chm5832703
f1_keywords:
- vbapb10.chm5832703
ms.prod: publisher
api_name:
- Publisher.InlineShapes
ms.assetid: 1a6d1e8f-0be0-102e-af6c-a1cee53eae02
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShapes Object (Publisher)

Contains a collection of  **[Shape](Publisher.Shape.md)** objects, which represent objects in the drawing layer, where **Shape.IsInline** is **True**. The collection of shapes is limited to shapes within a given text range.
 


## Remarks

The  **InlineShapes** collection is available only on the **TextRange** object. Using **TextFrame.Story.TextRange.InlineShapes** will return all inline shapes in a text frame, including those that are in overflow. Using **TextFrame.TextRange.InlineShapes** will return only visible inline shapes in a text frame, and not those that are in overflow.
 

 
The  **InlineShapes** collection can also be accessed from **Document.Stories( _i_ ).TextRange**, where i is the index to the active page of the publication.
 

 
The **InlineShapes** collection is not available in the **Page.Shapes** collection, including its contained **ShapeRange**.
 

 

## Example

Use the  **[InlineShapes](Publisher.TextRange.InlineShapes.md)** property on the **[TextRange](Publisher.TextRange.md)** object to return an **InlineShapes** collection. The following example finds the first shape, a text box, on page one of the publication, and appends text to the end of the text range in the text box if there is more than one inline shape within the text range.
 

 

```vb
Dim theShape As Shape 
 
Set theShape = ActiveDocument.Pages(1).Shapes(1) 
 
With theShape.TextFrame.TextRange 
 If .InlineShapes.Count > 1 Then 
 .InsertAfter (" There is more than one inline shape in this text box.") 
 End If 
End With
```

Use the  **InlineShapes** (index) property to return a single inline shape. The following example finds the third inline shape within a text box and flips it vertically.
 

 



```vb
Dim theShape As Shape 
 
Set theShape = ActiveDocument.Pages(1).Shapes(1) 
 
With theShape.TextFrame.Story.TextRange 
 With .InlineShapes(3) 
 .Flip (msoFlipVertical) 
 End With 
End With
```

Use the  **[Range](Publisher.Shapes.Range.md)** method to return a **[ShapeRange](Publisher.ShapeRange.md)** object that contains all members of the **InlineShapes** collection. An array of indexes or strings or a single index or string can be passed as a parameter of the **Range** property to select particular shapes or a shape within the range. The following example sets a **ShapeRange** variable equal to the collection of inline shapes that exist within a text box. Each inline shape within the range is then modified in some way. This example assumes that the first shape on the page is a text box that contains three inline shapes.
 

 



```vb
Dim theRange As ShapeRange 
 
Set theRange = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.Story.TextRange.InlineShapes.Range 
 
With theRange 
 .Item(1).Flip msoFlipVertical 
 .Item(2).MoveOutOfTextFlow 
 .Item(3).Delete 
End With
```


## Methods



|Name|
|:-----|
|[Item](Publisher.InlineShapes.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Publisher.InlineShapes.Application.md)|
|[Count](Publisher.InlineShapes.Count.md)|
|[Parent](Publisher.InlineShapes.Parent.md)|
|[Range](Publisher.InlineShapes.Range.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]