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

- [Apply](PowerPoint.Shape.Apply.md). Applies to the specified shape formatting that's been copied by using the **PickUp** method.
- [ApplyAnimation](PowerPoint.Shape.ApplyAnimation.md). Applies the last picked up animation to the **Shape** object.
- [ConvertTextToSmartArt](PowerPoint.Shape.ConvertTextToSmartArt.md). Converts text in a Shape object to a SmartArt diagram.
- [Copy](PowerPoint.Shape.Copy.md). Copies the specified object to the Clipboard.
- [Cut](PowerPoint.Shape.Cut.md). Deletes the specified object and places it on the Clipboard.
- [Delete](PowerPoint.Shape.Delete.md). Deletes the specified **Shape** object.
- [Duplicate](PowerPoint.Shape.Duplicate.md). Creates a duplicate of the specified **Shape** object, adds the new shape to the **Shapes** collection, and then returns a new **ShapeRange** object. The duplicated objects are placed at the end of the **Shapes** collection.
- [Flip](PowerPoint.Shape.Flip.md). Flips the specified shape around its horizontal or vertical axis.
- [IncrementLeft](PowerPoint.Shape.IncrementLeft.md). Moves the specified shape horizontally by the specified number of points.
- [IncrementRotation](PowerPoint.Shape.IncrementRotation.md). Changes the rotation of the specified shape around the z-axis by the specified number of degrees. Use the **Rotation** property to set the absolute rotation of the shape.
- [IncrementTop](PowerPoint.Shape.IncrementTop.md). Moves the specified shape vertically by the specified number of points.
- [PickUp](PowerPoint.Shape.PickUp.md). Copies the formatting of the specified shape. Use the **Apply** method to apply the copied formatting to another shape.
- [PickupAnimation](PowerPoint.Shape.PickupAnimation.md). Picks up all animation from the **Shape** object.
- [RerouteConnections](PowerPoint.Shape.RerouteConnections.md). Reroutes connectors so that they take the shortest possible path between the shapes they connect.
- [ScaleHeight](PowerPoint.Shape.ScaleHeight.md). Scales the height of the shape by a specified factor.
- [ScaleWidth](PowerPoint.Shape.ScaleWidth.md). Scales the width of the shape by a specified factor.
- [Select](PowerPoint.Shape.Select.md). Selects the specified object.
- [SetShapesDefaultProperties](PowerPoint.Shape.SetShapesDefaultProperties.md). Applies the formatting for the specified shape to the default shape.
- [Ungroup](PowerPoint.Shape.Ungroup.md). Ungroups any grouped shapes in the specified shape or range of shapes.
- [UpgradeMedia](PowerPoint.Shape.UpgradeMedia.md). Converts a legacy media object to an updated media object.
- [ZOrder](PowerPoint.Shape.ZOrder.md). Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).

## Properties

- [ActionSettings](PowerPoint.Shape.ActionSettings.md). Returns an [ActionSettings](PowerPoint.Shape.ActionSettings.md) object that contains information about what action occurs when the user clicks or moves the mouse over the specified shape or text range during a slide show. Read-only.
- [Adjustments](PowerPoint.Shape.Adjustments.md). Returns an [Adjustments](PowerPoint.Shape.Adjustments.md) object that contains adjustment values for all the adjustments in the specified shape. Applies to any Shape object that represents an AutoShape, WordArt, or a connector. Read-only.
- [AlternativeText](PowerPoint.Shape.AlternativeText.md). Returns or sets the alternative text associated with a shape in a Web presentation. Read/write.
- [AnimationSettings](PowerPoint.Shape.AnimationSettings.md). Returns an [AnimationSettings](PowerPoint.Shape.AnimationSettings.md) object that represents all the special effects you can apply to the animation of the specified shape. Read-only.
- [Application](PowerPoint.Shape.Application.md). Returns an [Application](PowerPoint.Shape.Application.md) object that represents the creator of the specified object.
- [AutoShapeType](PowerPoint.Shape.AutoShapeType.md). Returns or sets the shape type for the specified **Shape** object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write.
- [BackgroundStyle](PowerPoint.Shape.BackgroundStyle.md). Sets or returns the background style of the specified object. Read/write.
- [BlackWhiteMode](PowerPoint.Shape.BlackWhiteMode.md). Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/write.
- [Callout](PowerPoint.Shape.Callout.md). Returns a [CalloutFormat](PowerPoint.CalloutFormat.md) object that contains callout formatting properties for the specified shape. Applies to **Shape** or **ShapeRange** objects that represent line callouts. Read-only.
- [Chart](PowerPoint.Shape.Chart.md). Returns a **Chart** object of the current **Shape** object. Read-only.
- [Child](PowerPoint.Shape.Child.md). **MsoTrue** if the shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only.
- [ConnectionSiteCount](PowerPoint.Shape.ConnectionSiteCount.md). Returns the number of connection sites on the specified shape. Read-only.
- [Connector](PowerPoint.Shape.Connector.md). Determines whether the specified shape is a connector. Read-only.
- [ConnectorFormat](PowerPoint.Shape.ConnectorFormat.md). Returns a [ConnectorFormat](PowerPoint.ConnectorFormat.md) object that contains connector formatting properties. Applies to **Shape** or **ShapeRange** objects that represent connectors. Read-only.
- [Creator](PowerPoint.Shape.Creator.md). Returns a **Long** that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.
- [CustomerData](PowerPoint.Shape.CustomerData.md). Returns a **CustomerData** object. Read-only.
- [Decorative](PowerPoint.Shape.Decorative.md). Sets or returns the decorative flag for the specified object. Read/write.
- [Fill](PowerPoint.Shape.Fill.md). Returns a [FillFormat](PowerPoint.FillFormat.md) object that contains fill formatting properties for the specified shape. Read-only.
- [Glow](PowerPoint.Shape.Glow.md). Returns the glow format for the specified shape. Read-only.
- [GraphicStyle](PowerPoint.Shape.GraphicStyle.md). Returns or sets an [MsoGraphicStyleIndex](Office.MsoGraphicStyleIndex.md) constant that represents the style of an SVG graphic. Read/write.
- [GroupItems](PowerPoint.Shape.GroupItems.md). Returns a [GroupShapes](PowerPoint.GroupShapes.md) object that represents the individual shapes in the specified group. Use the **Item** method of the **GroupShapes** object to return a single shape from the group. Read-only.
- [HasChart](PowerPoint.Shape.HasChart.md). Returns whether the shape represented by the specified object contains a chart. Read-only.
- [HasInkXML](PowerPoint.shape.hasinkxml.md). Returns an [MsoTriState](Office.MsoTriState.md) enumeration value that indicates whether the specified shape contains ink XML that can be retrieved via the [Shape.InkXML](PowerPoint.Shape.InkXml.md) property. Read-only.
- [HasSmartArt](PowerPoint.Shape.HasSmartArt.md). Returns **True** if the current **Shape** object contains a SmartArt diagram. Read-only.
- [HasTable](PowerPoint.Shape.HasTable.md). Returns whether the specified shape is a table. Read-only.
- [HasTextFrame](PowerPoint.Shape.HasTextFrame.md). Returns whether the specified shape has a text frame. Read-only.
- [Height](PowerPoint.Shape.Height.md). Returns or sets the height of the specified object, in points. Read/write.
- [HorizontalFlip](PowerPoint.Shape.HorizontalFlip.md). Returns whether the specified shape is flipped around the horizontal axis. Read-only.
- [Id](PowerPoint.Shape.Id.md). Returns a **Long** that identifies the shape or range of shapes. Read-only.
- [InkXML](PowerPoint.shape.inkxml.md). Returns a **String** that contains the InkActionML associated with the specified shape. Read-only.
- [IsNarration](PowerPoint.shape.isnarration.md). Specifies whether the specified shape range contains a narration. Read/write.
- [Left](PowerPoint.Shape.Left.md). Returns or sets a **Single** that represents the distance in points from the left edge of the shape's bounding box to the left edge of the slide. Read/write.
- [Line](PowerPoint.Shape.Line.md). Returns a [LineFormat](PowerPoint.LineFormat.md) object that contains line formatting properties for the specified shape. (For a line, the **LineFormat** object represents the line itself; for a shape with a border, the **LineFormat** object represents the border.) Read-only.
- [LinkFormat](PowerPoint.Shape.LinkFormat.md). Returns a [LinkFormat](PowerPoint.LinkFormat.md) object that contains the properties that are unique to linked OLE objects. Read-only.
- [LockAspectRatio](PowerPoint.Shape.LockAspectRatio.md). Determines whether the specified shape retains its original proportions when you resize it. Read/write.
- [MediaFormat](PowerPoint.Shape.MediaFormat.md). Allows access to the new audio or video object. Read-only.
- [MediaType](PowerPoint.Shape.MediaType.md). Returns the OLE media type. Read-only.
- [Model3D](PowerPoint.Shape.Model3D.md). Returns a [Model3DFormat](PowerPoint.Model3dFormat.md) object that represents the 3D properties of a 3D model object. Read-only.
- [Name](PowerPoint.Shape.Name.md). Gets or sets the name of the **Shape**.
- [Nodes](PowerPoint.Shape.Nodes.md). Returns a [ShapeNodes](PowerPoint.ShapeNodes.md) collection that represents the geometric description of the specified shape. Applies to **Shape** objects that represent freeform drawings.
- [OLEFormat](PowerPoint.Shape.OLEFormat.md). Returns an [OLEFormat](PowerPoint.OLEFormat.md) object that contains OLE formatting properties for the specified shape. Applies to **Shape** or **ShapeRange** objects that represent OLE objects. Read-only.
- [Parent](PowerPoint.Shape.Parent.md). Returns the parent object for the specified object.
- [ParentGroup](PowerPoint.Shape.ParentGroup.md). Returns a **Shape** object that represents the common parent shape of a child shape or a range of child shapes.
- [PictureFormat](PowerPoint.Shape.PictureFormat.md). Returns a [PictureFormat](PowerPoint.PictureFormat.md) object that contains picture formatting properties for the specified shape. Read-only.
- [PlaceholderFormat](PowerPoint.Shape.PlaceholderFormat.md). Returns a [PlaceholderFormat](PowerPoint.PlaceholderFormat.md) object that contains the properties that are unique to placeholders. Read-only.
- [Reflection](PowerPoint.Shape.Reflection.md). Returns the reflection format for the specified shape. Read-only.
- [Rotation](PowerPoint.Shape.Rotation.md). Returns or sets the number of degrees the specified shape is rotated around the z-axis. Read/write.
- [Shadow](PowerPoint.Shape.Shadow.md). Returns a [ShadowFormat](PowerPoint.ShadowFormat.md) object that contains shadow formatting properties for the specified shape. Read-only.
- [ShapeStyle](PowerPoint.Shape.ShapeStyle.md). Sets or returns the shape style index for the specified object. Read/write.
- [SmartArt](PowerPoint.Shape.SmartArt.md). Returns a Microsoft Office [SmartArt](Office.SmartArt.md) object that represents the SmartArt diagram of the Shape object. Read-only.
- [SoftEdge](PowerPoint.Shape.SoftEdge.md). Returns the soft edge format for the specified shape. Read-only.
- [Table](PowerPoint.Shape.Table.md). Returns a [Table](PowerPoint.Table.md) object that represents a table in a shape or in a shape range. Read-only.
- [Tags](PowerPoint.Shape.Tags.md). Returns a [Tags](PowerPoint.Tags.md) object that represents the tags for the specified object. Read-only.
- [TextEffect](PowerPoint.Shape.TextEffect.md). Returns a [TextEffectFormat](PowerPoint.TextEffectFormat.md) object that contains text-effect formatting properties for the specified shape. Read-only.
- [TextFrame](PowerPoint.Shape.TextFrame.md). Returns a [TextFrame](PowerPoint.TextFrame.md) object that contains the alignment and anchoring properties for the specified shape or master text style.
- [TextFrame2](PowerPoint.Shape.TextFrame2.md). Returns the [TextFrame2](PowerPoint.TextFrame2.md) object associated with the specified **Shape** object that contains the alignment and anchoring properties for the specified shape. Read-only.
- [ThreeD](PowerPoint.Shape.ThreeD.md). Returns a [ThreeDFormat](PowerPoint.ThreeDFormat.md] object that contains 3D - effect formatting properties for the specified shape. Read-only.
- [Title](PowerPoint.Shape.Title.md). Returns a **Shape** object that represents the slide title. Read-only.
- [Top](PowerPoint.Shape.Top.md). Returns or sets a **Single** that represents the distance from the top edge of the shape's bounding box to the top edge of the document. Read/write.
- [Type](PowerPoint.Shape.Type.md). Represents the type of shape or shapes in a range of shapes. Read-only.
- [VerticalFlip](PowerPoint.Shape.VerticalFlip.md). Determines whether the specified shape is flipped around the vertical axis. Read-only.
- [Vertices](PowerPoint.Shape.Vertices.md). Returns the coordinates of the specified freeform drawing's vertices (and control points for Bézier curves) as a series of coordinate pairs. Read-only.
- [Visible](PowerPoint.Shape.Visible.md). Returns or sets the visibility of the specified object or the formatting applied to the specified object. Read/write.
- [Width](PowerPoint.Shape.Width.md). Returns or sets the width of the specified object, in points. Read/write.
- [ZOrderPosition](PowerPoint.Shape.ZOrderPosition.md). Returns the position of the specified shape in the z-order. Read-only.

## See also

- [PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
