---
title: ShapeRange object (PowerPoint)
keywords: vbapp10.chm548000
f1_keywords:
- vbapp10.chm548000
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange
ms.assetid: 0a194183-380e-ffb6-9336-b5bd311e917d
ms.date: 04/25/2019
localization_priority: Normal
---


# ShapeRange object (PowerPoint)

Represents a shape range, which is a set of shapes on a document. A shape range can contain as few as a single shape or as many as all the shapes on the document.


## Remarks

You can include whichever shapes you want—chosen from among all the shapes on the document or all the shapes in the selection—to construct a shape range. For example, you could construct a **ShapeRange** collection that contains the first three shapes on a document, all the selected shapes on a document, or all the freeforms on a document.

For an overview of how to work with either a single shape or with more than one shape at a time, see [Work with shapes (drawing objects)](../powerpoint/How-to/work-with-shapes-drawing-objects.md).

The following examples describe how to:

- Return a set of shapes that you specify by name or index number.
    
- Return all or some of the selected shapes on a document.
    

## Example

Use **Shapes.Range** (_index_), where _index_ is the name or index number of the shape or an array that contains either names or index numbers of shapes, to return a **ShapeRange** collection that represents a set of shapes on a document. You can use the **Array** function to construct an array of names or index numbers. The following example sets the fill pattern for shapes one and three on _myDocument_.

```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.Range(Array(1, 3)).Fill _

    .Patterned msoPatternHorizontalBrick
```

<br/>

The following example sets the fill pattern for the shapes named Oval 4 and Rectangle 5 on _myDocument_.

```vb
Set myDocument = ActivePresentation.Slides(1)

Set myRange = myDocument.Shapes _

    .Range(Array("Oval 4", "Rectangle 5"))

myRange.Fill.Patterned msoPatternHorizontalBrick
```

<br/>

Although you can use the **[Range](PowerPoint.Shapes.Range.md)** method to return any number of shapes or slides, it is simpler to use the **[Item](PowerPoint.ShapeRange.Item.md)** method if you want to return only a single member of the collection. For example, `Shapes(1)` is simpler than `Shapes.Range(1)`.

Use the **[ShapeRange](PowerPoint.Selection.ShapeRange.md)** property of the **Selection** object to return all the shapes in the selection. The following example sets the fill foreground color for all the shapes in the selection in window one, assuming that there's at least one shape in the selection.

```vb
Windows(1).Selection.ShapeRange.Fill.ForeColor _

    .RGB = RGB(255, 0, 255)
```

<br/>

Use **Selection.ShapeRange** (_index_), where _index_ is the shape name or the index number, to return a single shape within the selection. The following example sets the fill foreground color for shape two in the collection of selected shapes in window one, assuming that there are at least two shapes in the selection.

```vb
Windows(1).Selection.ShapeRange(2).Fill.ForeColor _

    .RGB = RGB(255, 0, 255)
```


## Methods

- [Align](PowerPoint.ShapeRange.Align.md)
- [Apply](PowerPoint.ShapeRange.Apply.md)
- [ApplyAnimation](PowerPoint.ShapeRange.ApplyAnimation.md)
- [ConvertTextToSmartArt](PowerPoint.ShapeRange.ConvertTextToSmartArt.md)
- [Copy](PowerPoint.ShapeRange.Copy.md)
- [Cut](PowerPoint.ShapeRange.Cut.md)
- [Delete](PowerPoint.ShapeRange.Delete.md)
- [Distribute](PowerPoint.ShapeRange.Distribute.md)
- [Duplicate](PowerPoint.ShapeRange.Duplicate.md)
- [Flip](PowerPoint.ShapeRange.Flip.md)
- [Group](PowerPoint.ShapeRange.Group.md)
- [IncrementLeft](PowerPoint.ShapeRange.IncrementLeft.md)
- [IncrementRotation](PowerPoint.ShapeRange.IncrementRotation.md)
- [IncrementTop](PowerPoint.ShapeRange.IncrementTop.md)
- [Item](PowerPoint.ShapeRange.Item.md)
- [MergeShapes](PowerPoint.shaperange.mergeshapes.md)
- [PickUp](PowerPoint.ShapeRange.PickUp.md)
- [PickupAnimation](PowerPoint.ShapeRange.PickupAnimation.md)
- [Regroup](PowerPoint.ShapeRange.Regroup.md)
- [RerouteConnections](PowerPoint.ShapeRange.RerouteConnections.md)
- [ScaleHeight](PowerPoint.ShapeRange.ScaleHeight.md)
- [ScaleWidth](PowerPoint.ShapeRange.ScaleWidth.md)
- [Select](PowerPoint.ShapeRange.Select.md)
- [SetShapesDefaultProperties](PowerPoint.ShapeRange.SetShapesDefaultProperties.md)
- [Ungroup](PowerPoint.ShapeRange.Ungroup.md)
- [UpgradeMedia](PowerPoint.ShapeRange.UpgradeMedia.md)
- [ZOrder](PowerPoint.ShapeRange.ZOrder.md)

## Properties

- [ActionSettings](PowerPoint.ShapeRange.ActionSettings.md)
- [Adjustments](PowerPoint.ShapeRange.Adjustments.md)
- [AlternativeText](PowerPoint.ShapeRange.AlternativeText.md)
- [AnimationSettings](PowerPoint.ShapeRange.AnimationSettings.md)
- [Application](PowerPoint.ShapeRange.Application.md)
- [AutoShapeType](PowerPoint.ShapeRange.AutoShapeType.md)
- [BackgroundStyle](PowerPoint.ShapeRange.BackgroundStyle.md)
- [BlackWhiteMode](PowerPoint.ShapeRange.BlackWhiteMode.md)
- [Callout](PowerPoint.ShapeRange.Callout.md)
- [Chart](PowerPoint.ShapeRange.Chart.md)
- [Child](PowerPoint.ShapeRange.Child.md)
- [ConnectionSiteCount](PowerPoint.ShapeRange.ConnectionSiteCount.md)
- [Connector](PowerPoint.ShapeRange.Connector.md)
- [ConnectorFormat](PowerPoint.ShapeRange.ConnectorFormat.md)
- [Count](PowerPoint.ShapeRange.Count.md)
- [Creator](PowerPoint.ShapeRange.Creator.md)
- [CustomerData](PowerPoint.ShapeRange.CustomerData.md)
- [Decorative](PowerPoint.ShapeRange.Decorative.md)
- [Fill](PowerPoint.ShapeRange.Fill.md)
- [Glow](PowerPoint.ShapeRange.Glow.md)
- [GraphicStyle](PowerPoint.ShapeRange.GraphicStyle.md)
- [GroupItems](PowerPoint.ShapeRange.GroupItems.md)
- [HasChart](PowerPoint.ShapeRange.HasChart.md)
- [HasInkXML](PowerPoint.shaperange.hasinkxml.md)
- [HasSmartArt](PowerPoint.ShapeRange.HasSmartArt.md)
- [HasTable](PowerPoint.ShapeRange.HasTable.md)
- [HasTextFrame](PowerPoint.ShapeRange.HasTextFrame.md)
- [Height](PowerPoint.ShapeRange.Height.md)
- [HorizontalFlip](PowerPoint.ShapeRange.HorizontalFlip.md)
- [Id](PowerPoint.ShapeRange.Id.md)
- [InkXML](PowerPoint.shaperange.inkxml.md)
- [IsNarration](PowerPoint.shaperange.isnarration.md)
- [Left](PowerPoint.ShapeRange.Left.md)
- [Line](PowerPoint.ShapeRange.Line.md)
- [LinkFormat](PowerPoint.ShapeRange.LinkFormat.md)
- [LockAspectRatio](PowerPoint.ShapeRange.LockAspectRatio.md)
- [MediaFormat](PowerPoint.ShapeRange.MediaFormat.md)
- [MediaType](PowerPoint.ShapeRange.MediaType.md)
- [Model3D](PowerPoint.ShapeRange.Model3D.md)
- [Name](PowerPoint.ShapeRange.Name.md)
- [Nodes](PowerPoint.ShapeRange.Nodes.md)
- [OLEFormat](PowerPoint.ShapeRange.OLEFormat.md)
- [Parent](PowerPoint.ShapeRange.Parent.md)
- [ParentGroup](PowerPoint.ShapeRange.ParentGroup.md)
- [PictureFormat](PowerPoint.ShapeRange.PictureFormat.md)
- [PlaceholderFormat](PowerPoint.ShapeRange.PlaceholderFormat.md)
- [Reflection](PowerPoint.ShapeRange.Reflection.md)
- [Rotation](PowerPoint.ShapeRange.Rotation.md)
- [Shadow](PowerPoint.ShapeRange.Shadow.md)
- [ShapeStyle](PowerPoint.ShapeRange.ShapeStyle.md)
- [SmartArt](PowerPoint.ShapeRange.SmartArt.md)
- [SoftEdge](PowerPoint.ShapeRange.SoftEdge.md)
- [Table](PowerPoint.ShapeRange.Table.md)
- [Tags](PowerPoint.ShapeRange.Tags.md)
- [TextEffect](PowerPoint.ShapeRange.TextEffect.md)
- [TextFrame](PowerPoint.ShapeRange.TextFrame.md)
- [TextFrame2](PowerPoint.ShapeRange.TextFrame2.md)
- [ThreeD](PowerPoint.ShapeRange.ThreeD.md)
- [Title](PowerPoint.ShapeRange.Title.md)
- [Top](PowerPoint.ShapeRange.Top.md)
- [Type](PowerPoint.ShapeRange.Type.md)
- [VerticalFlip](PowerPoint.ShapeRange.VerticalFlip.md)
- [Vertices](PowerPoint.ShapeRange.Vertices.md)
- [Visible](PowerPoint.ShapeRange.Visible.md)
- [Width](PowerPoint.ShapeRange.Width.md)
- [ZOrderPosition](PowerPoint.ShapeRange.ZOrderPosition.md)

## See also

- [PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
