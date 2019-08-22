---
title: Shape object (PowerPoint)
keywords: vbapp10.chm547000
f1_keywords:
- vbapp10.chm547000
ms.prod: powerpoint
api_name:
- PowerPoint.Shape
ms.assetid: 1da93849-99e0-827e-ced3-c6cf7f8569f3
ms.date: 04/25/2019
localization_priority: Normal
---


# Shape object (PowerPoint)

Represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.


## Remarks

> [!NOTE] 
> There are three objects that represent shapes: the **Shapes** collection, which represents all the shapes on a document; the **[ShapeRange](PowerPoint.ShapeRange.md)** collection, which represents a specified subset of the shapes on a document (for example, a **ShapeRange** object could represent shapes one and four on the document, or it could represent all the selected shapes on the document); and the **Shape** object, which represents a single shape on a document. If you want to work with several shapes at the same time or with shapes within the selection, use a **ShapeRange** collection. 
> 
> For an overview of how to work with either a single shape or with more than one shape at a time, see [Work with shapes (drawing objects)](../powerpoint/How-to/work-with-shapes-drawing-objects.md).

The following examples describe how to:

- Return an existing shape on a slide, indexed by name or number.
    
- Return a newly created shape on a slide.
    
- Return a shape within the selection.
    
- Return the slide title and other placeholders on a slide.
    
- Return the shapes attached to the ends of a connector.
    
- Return the default shape for a presentation.
    
- Return a newly created freeform.
    
- Return a single shape from within a group.
    
- Return a newly formed group of shapes.
    

## Example

Use **Shapes** (_index_), where _index_ is the shape name or the index number, to return a **Shape** object that represents a shape on a slide. The following example horizontally flips shape one and the shape named Rectangle 1 on _myDocument_.

```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(1).Flip msoFlipHorizontal

myDocument.Shapes("Rectangle 1").Flip msoFlipHorizontal
```

<br/>

Each shape is assigned a default name when you add it to the **Shapes** collection. To give the shape a more meaningful name, use the **Name** property. The following example adds a rectangle to _myDocument_, gives it the name Red Square, and then sets its foreground color and line style.

```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddShape(Type:=msoShapeRectangle, _

        Top:=144, Left:=144, Width:=72, Height:=72)

    .Name = "Red Square"

    .Fill.ForeColor.RGB = RGB(255, 0, 0)

    .Line.DashStyle = msoLineDashDot

End With
```

<br/>

To add a shape to a slide and return a **Shape** object that represents the newly created shape, use one of the following methods of the **Shapes** collection: [Add3DModel](PowerPoint.Shapes.Add3DModel.md), [AddCallout](PowerPoint.Shapes.AddCallout.md), [AddConnector](PowerPoint.Shapes.AddConnector.md), [AddCurve](PowerPoint.Shapes.AddCurve.md), [AddLabel](PowerPoint.Shapes.AddLabel.md), [AddLine](PowerPoint.Shapes.AddLine.md), [AddMediaObject](PowerPoint.Shapes.AddMediaObject.md), [AddOLEObject](PowerPoint.Shapes.AddOLEObject.md), [AddPicture](PowerPoint.Shapes.AddPicture.md), [AddPlaceholder](PowerPoint.Shapes.AddPlaceholder.md), [AddPolyline](PowerPoint.Shapes.AddPolyline.md), [AddShape](PowerPoint.Shapes.AddShape.md), [AddTable](PowerPoint.Shapes.AddTable.md), [AddTextbox](PowerPoint.Shapes.AddTextbox.md), [AddTextEffect](PowerPoint.Shapes.AddTextEffect.md), [AddTitle](PowerPoint.Shapes.AddTitle.md).

Use **Selection.ShapeRange** (_index_), where _index_ is the shape name or the index number, to return a **Shape** object that represents a shape within the selection. The following example sets the fill for the first shape in the selection in the active window, assuming that there's at least one shape in the selection.

```vb
ActiveWindow.Selection.ShapeRange(1).Fill _

    .ForeColor.RGB = RGB(255, 0, 0)
```

<br/>

Use **Shapes.Title** to return a **Shape** object that represents an existing slide title. Use **Shapes.AddTitle** to add a title to a slide that doesn't already have one and return a **Shape** object that represents the newly created title. Use **Shapes.Placeholders** (_index_), where _index_ is the placeholder's index number, to return a **Shape** object that represents a placeholder. If you have not changed the layering order of the shapes on a slide, the following three statements are equivalent, assuming that slide one has a title.

```vb
ActivePresentation.Slides(1).Shapes.Title _

    .TextFrame.TextRange.Font.Italic = True

ActivePresentation.Slides(1).Shapes.Placeholders(1) _

    .TextFrame.TextRange.Font.Italic = True

ActivePresentation.Slides(1).Shapes(1).TextFrame _

    .TextRange.Font.Italic = True
```

<br/>

To return a **Shape** object that represents one of the shapes attached by a connector, use the **[BeginConnectedShape](PowerPoint.ConnectorFormat.BeginConnectedShape.md)** or **[EndConnectedShape](PowerPoint.ConnectorFormat.EndConnectedShape.md)** property.

To return a **Shape** object that represents the default shape for a presentation, use the **[DefaultShape](PowerPoint.Presentation.DefaultShape.md)** property.

Use the **[BuildFreeform](PowerPoint.Shapes.BuildFreeform.md)** and **[AddNodes](PowerPoint.FreeformBuilder.AddNodes.md)** methods to define the geometry of a new freeform, and use the **[ConvertToShape](PowerPoint.FreeformBuilder.ConvertToShape.md)** method to create the freeform and return the **Shape** object that represents it.

Use **GroupItems** (_index_), where _index_ is the shape name or the index number within the group, to return a **Shape** object that represents a single shape in a grouped shape.

Use the **[Group](PowerPoint.ShapeRange.Group.md)** or **[Regroup](PowerPoint.ShapeRange.Regroup.md)** method to group a range of shapes and return a single **Shape** object that represents the newly formed group. After a group has been formed, you can work with the group the same way you work with any other shape.


## Methods

- [Apply](PowerPoint.Shape.Apply.md)
- [ApplyAnimation](PowerPoint.Shape.ApplyAnimation.md)
- [ConvertTextToSmartArt](PowerPoint.Shape.ConvertTextToSmartArt.md)
- [Copy](PowerPoint.Shape.Copy.md)
- [Cut](PowerPoint.Shape.Cut.md)
- [Delete](PowerPoint.Shape.Delete.md)
- [Duplicate](PowerPoint.Shape.Duplicate.md)
- [Flip](PowerPoint.Shape.Flip.md)
- [IncrementLeft](PowerPoint.Shape.IncrementLeft.md)
- [IncrementRotation](PowerPoint.Shape.IncrementRotation.md)
- [IncrementTop](PowerPoint.Shape.IncrementTop.md)
- [PickUp](PowerPoint.Shape.PickUp.md)
- [PickupAnimation](PowerPoint.Shape.PickupAnimation.md)
- [RerouteConnections](PowerPoint.Shape.RerouteConnections.md)
- [ScaleHeight](PowerPoint.Shape.ScaleHeight.md)
- [ScaleWidth](PowerPoint.Shape.ScaleWidth.md)
- [Select](PowerPoint.Shape.Select.md)
- [SetShapesDefaultProperties](PowerPoint.Shape.SetShapesDefaultProperties.md)
- [Ungroup](PowerPoint.Shape.Ungroup.md)
- [UpgradeMedia](PowerPoint.Shape.UpgradeMedia.md)
- [ZOrder](PowerPoint.Shape.ZOrder.md)

## Properties

- [ActionSettings](PowerPoint.Shape.ActionSettings.md)
- [Adjustments](PowerPoint.Shape.Adjustments.md)
- [AlternativeText](PowerPoint.Shape.AlternativeText.md)
- [AnimationSettings](PowerPoint.Shape.AnimationSettings.md)
- [Application](PowerPoint.Shape.Application.md)
- [AutoShapeType](PowerPoint.Shape.AutoShapeType.md)
- [BackgroundStyle](PowerPoint.Shape.BackgroundStyle.md)
- [BlackWhiteMode](PowerPoint.Shape.BlackWhiteMode.md)
- [Callout](PowerPoint.Shape.Callout.md)
- [Chart](PowerPoint.Shape.Chart.md)
- [Child](PowerPoint.Shape.Child.md)
- [ConnectionSiteCount](PowerPoint.Shape.ConnectionSiteCount.md)
- [Connector](PowerPoint.Shape.Connector.md)
- [ConnectorFormat](PowerPoint.Shape.ConnectorFormat.md)
- [Creator](PowerPoint.Shape.Creator.md)
- [CustomerData](PowerPoint.Shape.CustomerData.md)
- [Decorative](PowerPoint.Shape.Decorative.md)
- [Fill](PowerPoint.Shape.Fill.md)
- [Glow](PowerPoint.Shape.Glow.md)
- [GraphicStyle](PowerPoint.Shape.GraphicStyle.md)
- [GroupItems](PowerPoint.Shape.GroupItems.md)
- [HasChart](PowerPoint.Shape.HasChart.md)
- [HasInkXML](PowerPoint.shape.hasinkxml.md)
- [HasSmartArt](PowerPoint.Shape.HasSmartArt.md)
- [HasTable](PowerPoint.Shape.HasTable.md)
- [HasTextFrame](PowerPoint.Shape.HasTextFrame.md)
- [Height](PowerPoint.Shape.Height.md)
- [HorizontalFlip](PowerPoint.Shape.HorizontalFlip.md)
- [Id](PowerPoint.Shape.Id.md)
- [InkXML](PowerPoint.shape.inkxml.md)
- [IsNarration](PowerPoint.shape.isnarration.md)
- [Left](PowerPoint.Shape.Left.md)
- [Line](PowerPoint.Shape.Line.md)
- [LinkFormat](PowerPoint.Shape.LinkFormat.md)
- [LockAspectRatio](PowerPoint.Shape.LockAspectRatio.md)
- [MediaFormat](PowerPoint.Shape.MediaFormat.md)
- [MediaType](PowerPoint.Shape.MediaType.md)
- [Model3D](PowerPoint.Shape.Model3D.md)
- [Name](PowerPoint.Shape.Name.md)
- [Nodes](PowerPoint.Shape.Nodes.md)
- [OLEFormat](PowerPoint.Shape.OLEFormat.md)
- [Parent](PowerPoint.Shape.Parent.md)
- [ParentGroup](PowerPoint.Shape.ParentGroup.md)
- [PictureFormat](PowerPoint.Shape.PictureFormat.md)
- [PlaceholderFormat](PowerPoint.Shape.PlaceholderFormat.md)
- [Reflection](PowerPoint.Shape.Reflection.md)
- [Rotation](PowerPoint.Shape.Rotation.md)
- [Shadow](PowerPoint.Shape.Shadow.md)
- [ShapeStyle](PowerPoint.Shape.ShapeStyle.md)
- [SmartArt](PowerPoint.Shape.SmartArt.md)
- [SoftEdge](PowerPoint.Shape.SoftEdge.md)
- [Table](PowerPoint.Shape.Table.md)
- [Tags](PowerPoint.Shape.Tags.md)
- [TextEffect](PowerPoint.Shape.TextEffect.md)
- [TextFrame](PowerPoint.Shape.TextFrame.md)
- [TextFrame2](PowerPoint.Shape.TextFrame2.md)
- [ThreeD](PowerPoint.Shape.ThreeD.md)
- [Title](PowerPoint.Shape.Title.md)
- [Top](PowerPoint.Shape.Top.md)
- [Type](PowerPoint.Shape.Type.md)
- [VerticalFlip](PowerPoint.Shape.VerticalFlip.md)
- [Vertices](PowerPoint.Shape.Vertices.md)
- [Visible](PowerPoint.Shape.Visible.md)
- [Width](PowerPoint.Shape.Width.md)
- [ZOrderPosition](PowerPoint.Shape.ZOrderPosition.md)

## See also

- [PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
