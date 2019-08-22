---
title: ShapeRange object (Word)
keywords: vbawd10.chm2485
f1_keywords:
- vbawd10.chm2485
ms.prod: word
ms.assetid: 7112acc0-e241-16ef-77bc-101b72d05af0
ms.date: 04/25/2019
localization_priority: Normal
---


# ShapeRange object (Word)

Represents a shape range, which is a set of shapes on a document. A shape range can contain as few as one shape or as many as all the shapes in the document. 


## Remarks

You can include whichever shapes you want—chosen from among all the shapes in the document or all the shapes in the selection—to construct a shape range. For example, you could construct a **ShapeRange** collection that contains the first three shapes in a document, all the selected shapes in a document, or all the freeform shapes in a document. Most operations that you can do with a **Shape** object, you can also do with a **ShapeRange** object that contains only one shape. Some operations, when performed on a **ShapeRange** object that contains more than one shape, will cause an error.

Use **Range** (_index_), where _index_ is the name or index number of the shape or an array that contains either names or index numbers of shapes, to return a **ShapeRange** collection that represents a set of shapes on a document. You can use Visual Basic's **Array** function to construct an array of names or index numbers. The following example sets the fill pattern for shapes one and three on the active document.

```vb
ActiveDocument.Shapes.Range(Array(1, 3)).Fill.Patterned _ 
 msoPatternHorizontalBrick
```

<br/>

The following example selects the shapes named Oval 4 and Rectangle 5 on the active document.

```vb
ActiveDocument.Shapes.Range(Array("Oval 4", "Rectangle 5")).Select
```

<br/>

Although you can use the **Range** method to return any number of shapes, it is simpler to use the **Item** method if you want to return only a single member of the collection. For example, `Shapes(1)` is simpler than `Shapes.Range(1)`.

Use **ShapeRange** (_index_), where _index_ is the name or the index number, to return a **Shape** object that represents a shape within a selection. The following example sets the fill for the first shape in the selection, assuming that the selection contains at least one shape.

```vb
Selection.ShapeRange(1).Fill.ForeColor.RGB = RGB(255, 0, 0)
```

<br/>

This example selects all the shapes in the first section of the active document.

```vb
Set myRange = ActiveDocument.Sections(1).Range 
myRange.ShapeRange.Select
```

Use the **Align**, **Distribute**, or **ZOrder** method to position a set of shapes relative to each other or relative to the document.

Use the **Group**, **Regroup**, or **UnGroup** method to create and work with a single shape formed from a shape range. The **GroupItems** property for a **Shape** object returns the **GroupShapes** object, which represents all the shapes that were grouped to form one shape.

The recorder always uses the **ShapeRange** property when recording shapes.

> [!NOTE] 
> A **ShapeRange** object doesn't include **InlineShape** objects.


## Methods

- [Align](Word.ShapeRange.Align.md)
- [Apply](Word.ShapeRange.Apply.md)
- [CanvasCropBottom](Word.ShapeRange.CanvasCropBottom.md)
- [CanvasCropLeft](Word.ShapeRange.CanvasCropLeft.md)
- [CanvasCropRight](Word.ShapeRange.CanvasCropRight.md)
- [CanvasCropTop](Word.ShapeRange.CanvasCropTop.md)
- [ConvertToInlineShape](Word.ShapeRange.ConvertToInlineShape.md)
- [Delete](Word.ShapeRange.Delete.md)
- [Distribute](Word.ShapeRange.Distribute.md)
- [Duplicate](Word.ShapeRange.Duplicate.md)
- [Flip](Word.ShapeRange.Flip.md)
- [Group](Word.ShapeRange.Group.md)
- [IncrementLeft](Word.ShapeRange.IncrementLeft.md)
- [IncrementRotation](Word.ShapeRange.IncrementRotation.md)
- [IncrementTop](Word.ShapeRange.IncrementTop.md)
- [Item](Word.ShapeRange.Item.md)
- [PickUp](Word.ShapeRange.PickUp.md)
- [ScaleHeight](Word.ShapeRange.ScaleHeight.md)
- [ScaleWidth](Word.ShapeRange.ScaleWidth.md)
- [Select](Word.ShapeRange.Select.md)
- [SetShapesDefaultProperties](Word.ShapeRange.SetShapesDefaultProperties.md)
- [Ungroup](Word.ShapeRange.Ungroup.md)
- [ZOrder](Word.ShapeRange.ZOrder.md)

## Properties

- [Adjustments](Word.ShapeRange.Adjustments.md)
- [AlternativeText](Word.ShapeRange.AlternativeText.md)
- [Anchor](Word.ShapeRange.Anchor.md)
- [Application](Word.ShapeRange.Application.md)
- [AutoShapeType](Word.ShapeRange.AutoShapeType.md)
- [BackgroundStyle](Word.ShapeRange.BackgroundStyle.md)
- [Callout](Word.ShapeRange.Callout.md)
- [CanvasItems](Word.ShapeRange.CanvasItems.md)
- [Child](Word.ShapeRange.Child.md)
- [Count](Word.ShapeRange.Count.md)
- [Creator](Word.ShapeRange.Creator.md)
- [Decorative](Word.ShapeRange.Decorative.md)
- [Fill](Word.ShapeRange.Fill.md)
- [Glow](Word.ShapeRange.Glow.md)
- [GraphicStyle](Word.ShapeRange.GraphicStyle.md)
- [GroupItems](Word.ShapeRange.GroupItems.md)
- [Height](Word.ShapeRange.Height.md)
- [HeightRelative](Word.ShapeRange.HeightRelative.md)
- [HorizontalFlip](Word.ShapeRange.HorizontalFlip.md)
- [Hyperlink](Word.ShapeRange.Hyperlink.md)
- [ID](Word.ShapeRange.ID.md)
- [LayoutInCell](Word.ShapeRange.LayoutInCell.md)
- [Left](Word.ShapeRange.Left.md)
- [LeftRelative](Word.ShapeRange.LeftRelative.md)
- [Line](Word.ShapeRange.Line.md)
- [LockAnchor](Word.ShapeRange.LockAnchor.md)
- [LockAspectRatio](Word.ShapeRange.LockAspectRatio.md)
- [Model3D](Word.ShapeRange.Model3D.md)
- [Name](Word.ShapeRange.Name.md)
- [Nodes](Word.ShapeRange.Nodes.md)
- [Parent](Word.ShapeRange.Parent.md)
- [ParentGroup](Word.ShapeRange.ParentGroup.md)
- [PictureFormat](Word.ShapeRange.PictureFormat.md)
- [Reflection](Word.ShapeRange.Reflection.md)
- [RelativeHorizontalPosition](Word.ShapeRange.RelativeHorizontalPosition.md)
- [RelativeHorizontalSize](Word.ShapeRange.RelativeHorizontalSize.md)
- [RelativeVerticalPosition](Word.ShapeRange.RelativeVerticalPosition.md)
- [RelativeVerticalSize](Word.ShapeRange.RelativeVerticalSize.md)
- [Rotation](Word.ShapeRange.Rotation.md)
- [Shadow](Word.ShapeRange.Shadow.md)
- [ShapeStyle](Word.ShapeRange.ShapeStyle.md)
- [SoftEdge](Word.ShapeRange.SoftEdge.md)
- [TextEffect](Word.ShapeRange.TextEffect.md)
- [TextFrame](Word.ShapeRange.TextFrame.md)
- [TextFrame2](Word.ShapeRange.TextFrame2.md)
- [ThreeD](Word.ShapeRange.ThreeD.md)
- [Title](Word.ShapeRange.Title.md)
- [Top](Word.ShapeRange.Top.md)
- [TopRelative](Word.ShapeRange.TopRelative.md)
- [Type](Word.ShapeRange.Type.md)
- [VerticalFlip](Word.ShapeRange.VerticalFlip.md)
- [Vertices](Word.ShapeRange.Vertices.md)
- [Visible](Word.ShapeRange.Visible.md)
- [Width](Word.ShapeRange.Width.md)
- [WidthRelative](Word.ShapeRange.WidthRelative.md)
- [WrapFormat](Word.ShapeRange.WrapFormat.md)
- [ZOrderPosition](Word.ShapeRange.ZOrderPosition.md)

## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
