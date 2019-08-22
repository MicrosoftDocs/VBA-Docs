---
title: ShapeRange object (Excel)
keywords: vbaxl10.chm639072
f1_keywords:
- vbaxl10.chm639072
ms.prod: excel
api_name:
- Excel.ShapeRange
ms.assetid: e1b8229c-73a0-4a77-5e00-4bcec9032260
ms.date: 04/25/2019
localization_priority: Normal
---


# ShapeRange object (Excel)

Represents a shape range, which is a set of shapes on a document.


## Remarks

A shape range can contain as few as a single shape or as many as all the shapes on the document. You can include whichever shapes you want—chosen from among all the shapes on the document or all the shapes in the selection—to construct a shape range. For example, you could construct a **ShapeRange** collection that contains the first three shapes on a document, all the selected shapes on a document, or all the freeforms on a document.

## Example

### Return a set of shapes that you specify by name or index number

Use **[Range](excel.shapes.range.md)** (_index_), where _index_ is the name or index number of the shape or an array that contains either names or index numbers of shapes, to return a **ShapeRange** collection that represents a set of shapes on a document. You can use the **Array** function to construct an array of names or index numbers. 

The following example sets the fill pattern for shapes one and three on _myDocument_.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.Range(Array(1, 3)).Fill.Patterned _ 
 msoPatternHorizontalBrick
```

<br/>

The following example sets the fill pattern for the shapes named Oval 4 and Rectangle 5 on _myDocument_.

Although you can use the **Range** property to return any number of shapes or slides, it's simpler to use the **Item** method if you want to return only a single member of the collection. For example, `Shapes(1)` is simpler than `Shapes.Range(1)`.

```vb
Set myDocument = Worksheets(1) 
Set myRange = myDocument.Shapes.Range(Array("Oval 4", _ 
 "Rectangle 5")) 
myRange.Fill.Patterned msoPatternHorizontalBrick
```

### Return all or some of the selected shapes on a document

Use the **ShapeRange** property of the **Selection** object to return all the shapes in the selection. The following example sets the fill foreground color for all the shapes in the selection in window one, assuming that there's at least one shape in the selection.

```vb
Windows(1).Selection.ShapeRange.Fill.ForeColor.RGB = _ 
 RGB(255, 0, 255)
```

<br/>

Use _Selection_.**ShapeRange** (_index_), where _index_ is the shape name or the index number, to return a single shape within the selection. The following example sets the fill foreground color for shape two in the collection of selected shapes in window one, assuming that there are at least two shapes in the selection.

```vb
Windows(1).Selection.ShapeRange(2).Fill.ForeColor.RGB = _ 
 RGB(255, 0, 255)
```


## Methods

- [Align](Excel.ShapeRange.Align.md)
- [Apply](Excel.ShapeRange.Apply.md)
- [Delete](Excel.ShapeRange.Delete.md)
- [Distribute](Excel.ShapeRange.Distribute.md)
- [Duplicate](Excel.ShapeRange.Duplicate.md)
- [Flip](Excel.ShapeRange.Flip.md)
- [Group](Excel.ShapeRange.Group.md)
- [IncrementLeft](Excel.ShapeRange.IncrementLeft.md)
- [IncrementRotation](Excel.ShapeRange.IncrementRotation.md)
- [IncrementTop](Excel.ShapeRange.IncrementTop.md)
- [Item](Excel.ShapeRange.Item.md)
- [PickUp](Excel.ShapeRange.PickUp.md)
- [Regroup](Excel.ShapeRange.Regroup.md)
- [RerouteConnections](Excel.ShapeRange.RerouteConnections.md)
- [ScaleHeight](Excel.ShapeRange.ScaleHeight.md)
- [ScaleWidth](Excel.ShapeRange.ScaleWidth.md)
- [Select](Excel.ShapeRange.Select.md)
- [SetShapesDefaultProperties](Excel.ShapeRange.SetShapesDefaultProperties.md)
- [Ungroup](Excel.ShapeRange.Ungroup.md)
- [ZOrder](Excel.ShapeRange.ZOrder.md)

## Properties

- [Adjustments](Excel.ShapeRange.Adjustments.md)
- [AlternativeText](Excel.ShapeRange.AlternativeText.md)
- [Application](Excel.ShapeRange.Application.md)
- [AutoShapeType](Excel.ShapeRange.AutoShapeType.md)
- [BackgroundStyle](Excel.ShapeRange.BackgroundStyle.md)
- [BlackWhiteMode](Excel.ShapeRange.BlackWhiteMode.md)
- [Callout](Excel.ShapeRange.Callout.md)
- [Chart](Excel.ShapeRange.Chart.md)
- [Child](Excel.ShapeRange.Child.md)
- [ConnectionSiteCount](Excel.ShapeRange.ConnectionSiteCount.md)
- [Connector](Excel.ShapeRange.Connector.md)
- [ConnectorFormat](Excel.ShapeRange.ConnectorFormat.md)
- [Count](Excel.ShapeRange.Count.md)
- [Creator](Excel.ShapeRange.Creator.md)
- [Decorative](Excel.ShapeRange.Decorative.md)
- [Fill](Excel.ShapeRange.Fill.md)
- [Glow](Excel.ShapeRange.Glow.md)
- [GraphicStyle](Excel.ShapeRange.GraphicStyle.md)
- [GroupItems](Excel.ShapeRange.GroupItems.md)
- [HasChart](Excel.ShapeRange.HasChart.md)
- [Height](Excel.ShapeRange.Height.md)
- [HorizontalFlip](Excel.ShapeRange.HorizontalFlip.md)
- [ID](Excel.ShapeRange.ID.md)
- [Left](Excel.ShapeRange.Left.md)
- [Line](Excel.ShapeRange.Line.md)
- [LockAspectRatio](Excel.ShapeRange.LockAspectRatio.md)
- [Model3D](Excel.ShapeRange.Model3D.md)
- [Name](Excel.ShapeRange.Name.md)
- [Nodes](Excel.ShapeRange.Nodes.md)
- [Parent](Excel.ShapeRange.Parent.md)
- [ParentGroup](Excel.ShapeRange.ParentGroup.md)
- [PictureFormat](Excel.ShapeRange.PictureFormat.md)
- [Reflection](Excel.ShapeRange.Reflection.md)
- [Rotation](Excel.ShapeRange.Rotation.md)
- [Shadow](Excel.ShapeRange.Shadow.md)
- [ShapeStyle](Excel.ShapeRange.ShapeStyle.md)
- [SoftEdge](Excel.ShapeRange.SoftEdge.md)
- [TextEffect](Excel.ShapeRange.TextEffect.md)
- [TextFrame](Excel.ShapeRange.TextFrame.md)
- [TextFrame2](Excel.ShapeRange.TextFrame2.md)
- [ThreeD](Excel.ShapeRange.ThreeD.md)
- [Title](Excel.ShapeRange.Title.md)
- [Top](Excel.ShapeRange.Top.md)
- [Type](Excel.ShapeRange.Type.md)
- [VerticalFlip](Excel.ShapeRange.VerticalFlip.md)
- [Vertices](Excel.ShapeRange.Vertices.md)
- [Visible](Excel.ShapeRange.Visible.md)
- [Width](Excel.ShapeRange.Width.md)
- [ZOrderPosition](Excel.ShapeRange.ZOrderPosition.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
