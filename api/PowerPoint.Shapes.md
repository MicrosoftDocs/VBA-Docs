---
title: Shapes object (PowerPoint)
keywords: vbapp10.chm543000
f1_keywords:
- vbapp10.chm543000
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes
ms.assetid: eb208855-254e-1a0f-884b-4a5edcfd584d
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes object (PowerPoint)

A collection of all the  **[Shape](PowerPoint.Shape.md)** objects on the specified slide.


## Remarks

Each  **Shape** object represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.


> [!NOTE] 
> If you want to work with a subset of the shapes on a document — for example, to do something to only the AutoShapes on the document or to only the selected shapes — you must construct a **[ShapeRange](PowerPoint.ShapeRange.md)** collection that contains the shapes you want to work with. For an overview of how to work either with a single shape or with more than one shape at a time, see [How to: Work with Shapes (Drawing Objects)](../powerpoint/How-to/work-with-shapes-drawing-objects.md).


## Example

Use the  **Shapes** property to return the **Shapes** collection. The following example selects all the shapes in the active presentation.


```vb
ActivePresentation.Slides(1).Shapes.SelectAll
```


> [!NOTE] 
> If you want to do something (like delete or set a property) to all the shapes on a document at the same time, use the [Range](PowerPoint.Shapes.Range.md)method with no argument to create a **ShapeRange** object that contains all the shapes in the **Shapes** collection, and then apply the appropriate property or method to the **ShapeRange** object.

Use the [AddCallout](PowerPoint.Shapes.AddCallout.md), [AddComment](overview/PowerPoint.md), [AddConnector](PowerPoint.Shapes.AddConnector.md), [AddCurve](PowerPoint.Shapes.AddCurve.md), [AddLabel](PowerPoint.Shapes.AddLabel.md), [AddLine](PowerPoint.Shapes.AddLine.md), [AddMediaObject](PowerPoint.Shapes.AddMediaObject.md), [AddOLEObject](PowerPoint.Shapes.AddOLEObject.md), [AddPicture](PowerPoint.Shapes.AddPicture.md), [AddPlaceholder](PowerPoint.Shapes.AddPlaceholder.md), [AddPolyline](PowerPoint.Shapes.AddPolyline.md), [AddShape](PowerPoint.Shapes.AddShape.md), [AddTable](PowerPoint.Shapes.AddTable.md), [AddTextbox](PowerPoint.Shapes.AddTextbox.md), [AddTextEffect](PowerPoint.Shapes.AddTextEffect.md), or [AddTitle](PowerPoint.Shapes.AddTitle.md)method to create a new shape and add it to the  **Shapes** collection. Use the [BuildFreeform](PowerPoint.Shapes.BuildFreeform.md)method in conjunction with the [ConvertToShape](PowerPoint.FreeformBuilder.ConvertToShape.md)method to create a new freeform and add it to the collection. The following example adds a rectangle to the active presentation.




```vb
ActivePresentation.Slides(1).Shapes.AddShape Type:=msoShapeRectangle, _

    Left:=50, Top:=50, Width:=100, Height:=200
```

Use  **Shapes** (_index_), where _index_ is the shape's name or index number, to return a single **Shape** object. The following example sets the fill to a preset shade for shape one in the active presentation.




```vb
ActivePresentation.Slides(1).Shapes(1).Fill _

    .PresetGradient Style:=msoGradientHorizontal, Variant:=1, _

    PresetGradientType:=msoGradientBrass
```

Use  **Shapes.Range** (_index_), where _index_ is the shape's name or index number or an array of shape names or index numbers, to return a **[ShapeRange](PowerPoint.ShapeRange.md)** collection that represents a subset of the **Shapes** collection. The following example sets the fill pattern for shapes one and three in the active presentation.




```vb
ActivePresentation.Slides(1).Shapes.Range(Array(1, 3)).Fill _

    .Patterned Pattern:=msoPatternHorizontalBrick
```

Use  **Shapes.Placeholders** (_index_), where _index_ is the placeholder number, to return a **Shape** object that represents a placeholder. If the specified slide has a title, use **Shapes.Placeholders(1)** or **Shapes.Title** to return the title placeholder. The following example adds a slide to the active presentation and then adds text to both the title and the subtitle (the subtitle is the second placeholder on a slide with this layout).




```vb
With ActivePresentation.Slides.Add(Index:=1, Layout:=ppLayoutTitle).Shapes

    .Title.TextFrame.TextRange = "This is the title text"

    .Placeholders(2).TextFrame.TextRange = "This is subtitle text"

End With
```


## Methods



|Name|
|:-----|
|[AddCallout](PowerPoint.Shapes.AddCallout.md)|
|[AddChart2](PowerPoint.shapes.addchart2.md)|
|[AddConnector](PowerPoint.Shapes.AddConnector.md)|
|[AddCurve](PowerPoint.Shapes.AddCurve.md)|
|[AddInkShapeFromXML](PowerPoint.shapes.addinkshapefromxml.md)|
|[AddLabel](PowerPoint.Shapes.AddLabel.md)|
|[AddLine](PowerPoint.Shapes.AddLine.md)|
|[AddMediaObject2](PowerPoint.Shapes.AddMediaObject2.md)|
|[AddMediaObjectFromEmbedTag](PowerPoint.Shapes.AddMediaObjectFromEmbedTag.md)|
|[AddOLEObject](PowerPoint.Shapes.AddOLEObject.md)|
|[AddPicture](PowerPoint.Shapes.AddPicture.md)|
|[AddPicture2](PowerPoint.shapes.addpicture2.md)|
|[AddPlaceholder](PowerPoint.Shapes.AddPlaceholder.md)|
|[AddPolyline](PowerPoint.Shapes.AddPolyline.md)|
|[AddShape](PowerPoint.Shapes.AddShape.md)|
|[AddSmartArt](PowerPoint.Shapes.AddSmartArt.md)|
|[AddTable](PowerPoint.Shapes.AddTable.md)|
|[AddTextbox](PowerPoint.Shapes.AddTextbox.md)|
|[AddTextEffect](PowerPoint.Shapes.AddTextEffect.md)|
|[Add3DModel](PowerPoint.Shapes.Add3DModel.md)|
|[AddTitle](PowerPoint.Shapes.AddTitle.md)|
|[BuildFreeform](PowerPoint.Shapes.BuildFreeform.md)|
|[Item](PowerPoint.Shapes.Item.md)|
|[Paste](PowerPoint.Shapes.Paste.md)|
|[PasteSpecial](PowerPoint.Shapes.PasteSpecial.md)|
|[Range](PowerPoint.Shapes.Range.md)|
|[SelectAll](PowerPoint.Shapes.SelectAll.md)|

## Properties



|Name|
|:-----|
|[Application](PowerPoint.Shapes.Application.md)|
|[Count](PowerPoint.Shapes.Count.md)|
|[Creator](PowerPoint.Shapes.Creator.md)|
|[HasTitle](PowerPoint.Shapes.HasTitle.md)|
|[Parent](PowerPoint.Shapes.Parent.md)|
|[Placeholders](PowerPoint.Shapes.Placeholders.md)|
|[Title](PowerPoint.Shapes.Title.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
