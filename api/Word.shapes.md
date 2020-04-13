---
title: Shapes object (Word)
keywords: vbawd10.chm2463
f1_keywords:
- vbawd10.chm2463
ms.prod: word
ms.assetid: 0907eed3-886e-8e73-0e5e-71f4b37ddd5b
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes object (Word)

A collection of  **Shape** objects that represent all the shapes in a document or all the shapes in all the headers and footers in a document. Each **[Shape](Word.Shape.md)** object represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.


## Remarks

If you want to work with a subset of the shapes on a document — for example, to do something to only the AutoShapes on the document or to only the selected shapes — you must construct a  **[ShapeRange](Word.shaperange.md)** collection that contains the shapes you want to work with.

Use the **Shapes** property to return the **Shapes** collection. The following example selects all the shapes on the active document.




```vb
ActiveDocument.Shapes.SelectAll
```


> [!NOTE] 
> If you want to do something (like delete or set a property) to all the shapes on a document at the same time, use the **Range** method to create a **ShapeRange** object that contains all the shapes in the **Shapes** collection, and then apply the appropriate property or method to the **ShapeRange** object.

Use one of the following methods of the **Shapes** collection: **Add3DModel**, **AddCallout**, **AddCurve**, **AddLabel**, **AddLine**, **AddOleControl**, **AddOleObject**, **AddPolyline**, **AddShape**, **AddTextbox**, **AddTextEffect**, or **BuildFreeForm** to add a shape to a document return a **Shape** object that represents the newly created shape. The following example adds a rectangle to the active document.




```vb
ActiveDocument.Shapes.AddShape msoShapeRectangle, 50, 50, 100, 200
```

Use  **Shapes** (Index), where Index is the name or the index number, to return a single **Shape** object. The following example horizontally flips shape one on the active document.




```vb
ActiveDocument.Shapes(1).Flip msoFlipHorizontal
```

This example horizontally flips the shape named "Rectangle 1" on the active document.




```vb
ActiveDocument.Shapes("Rectangle 1").Flip msoFlipHorizontal
```

Each shape is assigned a default name when it is created. For example, if you add three different shapes to a document, they might be named "Rectangle 2," "TextBox 3," and "Oval 4." To give a shape a more meaningful name, set the **Name** property.

The **Shapes** collection does not include **[InlineShape](Word.InlineShape.md)** objects. **InlineShape** objects are treated like characters and are positioned as characters within a line of text. **Shape** objects are anchored to a range of text but are free-floating and can be positioned anywhere on the page. You can use the **ConvertToInlineShape** method and the **ConvertToShape** method to convert shapes from one type to the other. You can convert only pictures, OLE objects, and ActiveX controls to inline shapes.

The **Count** property for this collection in a document returns the number of items in the main story only. To count the shapes in all the headers and footers, use the **Shapes** collection with any **HeaderFooter** object.


## Methods



|Name|
|:-----|
|[AddCallout](Word.Shapes.AddCallout.md)|
|[AddCanvas](Word.Shapes.AddCanvas.md)|
|[AddChart2](Word.shapes.addchart2.md)|
|[AddCurve](Word.Shapes.AddCurve.md)|
|[AddLabel](Word.Shapes.AddLabel.md)|
|[AddLine](Word.Shapes.AddLine.md)|
|[AddOLEControl](Word.Shapes.AddOLEControl.md)|
|[AddOLEObject](Word.Shapes.AddOLEObject.md)|
|[AddPicture](Word.Shapes.AddPicture.md)|
|[AddPolyline](Word.Shapes.AddPolyline.md)|
|[AddShape](Word.Shapes.AddShape.md)|
|[AddSmartArt](Word.Shapes.AddSmartArt.md)|
|[AddTextbox](Word.Shapes.AddTextbox.md)|
|[AddTextEffect](Word.Shapes.AddTextEffect.md)|
|[Add3DModel](Word.Shapes.Add3DModel.md)|
|[AddWebVideo](Word.shapes.addwebvideo.md)|
|[BuildFreeform](Word.Shapes.BuildFreeform.md)|
|[Item](Word.Shapes.Item.md)|
|[Range](Word.Shapes.Range.md)|
|[SelectAll](Word.Shapes.SelectAll.md)|

## Properties



|Name|
|:-----|
|[Application](Word.Shapes.Application.md)|
|[Count](Word.Shapes.Count.md)|
|[Creator](Word.Shapes.Creator.md)|
|[Parent](Word.Shapes.Parent.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
