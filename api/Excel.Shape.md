---
title: Shape object (Excel)
keywords: vbaxl10.chm635072
f1_keywords:
- vbaxl10.chm635072
ms.prod: excel
api_name:
- Excel.Shape
ms.assetid: 8f01fcd1-b7d9-5216-2de5-40fb6648a403
ms.date: 04/25/2019
localization_priority: Normal
---


# Shape object (Excel)

Represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.


## Remarks

The **Shape** object is a member of the **[Shapes](Excel.Shapes.md)** collection. The **Shapes** collection contains all the shapes in a workbook.

> [!NOTE] 
> There are three objects that represent shapes: the **Shapes** collection, which represents all the shapes on a workbook; the **[ShapeRange](Excel.ShapeRange.md)** collection, which represents a specified subset of the shapes on a workbook (for example, a **ShapeRange** object could represent shapes one and four in the workbook, or it could represent all the selected shapes in the workbook); and the **Shape** object, which represents a single shape on a worksheet. If you want to work with several shapes at the same time or with shapes within the selection, use a **ShapeRange** collection.

|To return...|Use...|
|:-----------|:-----|
|A **Shape** object that represents one of the shapes attached by a connector |The **[BeginConnectedShape](Excel.ConnectorFormat.BeginConnectedShape.md)** or **[EndConnectedShape](Excel.ConnectorFormat.EndConnectedShape.md)** property of the **ConnectorFormat** object.|
|A newly created freeform |The **[BuildFreeform](Excel.Shapes.BuildFreeform.md)** and **[AddNodes](Excel.FreeformBuilder.AddNodes.md)** methods to define the geometry of a new freeform, and use the **[ConvertToShape](Excel.FreeformBuilder.ConvertToShape.md)** method to create the freeform and return the **Shape** object that represents it.|
|A **Shape** object that represents a single shape in a grouped shape |**[GroupItems](Excel.Shape.GroupItems.md)** (_index_), where _index_ is the shape name or the index number within the group.|
|A newly formed group of shapes |The **[Group](Excel.ShapeRange.Group.md)** or **[Regroup](Excel.ShapeRange.Regroup.md)** method of the **ShapeRange** object to group a range of shapes and return a single **Shape** object that represents the newly formed group. After a group has been formed, you can work with the group the same way that you work with any other shape.|
|A **Shape** object that represents an existing shape |**[Shapes](Excel.Worksheet.Shapes.md)** (_index_), where _index_ is the shape name or the index number.|
|A **Shape** object that represents a shape within the selection |_Selection_.**ShapeRange** (_index_), where _index_ is the shape name or the index number.|


## Example

The following example horizontally flips shape one and the shape named Rectangle 1 on _myDocument_.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).Flip msoFlipHorizontal 
myDocument.Shapes("Rectangle 1").Flip msoFlipHorizontal
```

<br/>

Each shape is assigned a default name when you add it to the **Shapes** collection. To give the shape a more meaningful name, use the **Name** property. The following example adds a rectangle to _myDocument_, gives it the name Red Square, and then sets its foreground color and line style.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 144, 144, 72, 72) 
 .Name = "Red Square" 
 .Fill.ForeColor.RGB = RGB(255, 0, 0) 
 .Line.DashStyle = msoLineDashDot 
End With
```

<br/>

The following example sets the fill for the first shape in the selection in the active window, assuming that there's at least one shape in the selection.

```vb
ActiveWindow.Selection.ShapeRange(1).Fill.ForeColor.RGB = _ 
 RGB(255, 0, 0)
```


## Methods

- [Apply](Excel.Shape.Apply.md)
- [Copy](Excel.Shape.Copy.md)
- [CopyPicture](Excel.Shape.CopyPicture.md)
- [Cut](Excel.Shape.Cut.md)
- [Delete](Excel.Shape.Delete.md)
- [Duplicate](Excel.Shape.Duplicate.md)
- [Flip](Excel.Shape.Flip.md)
- [IncrementLeft](Excel.Shape.IncrementLeft.md)
- [IncrementRotation](Excel.Shape.IncrementRotation.md)
- [IncrementTop](Excel.Shape.IncrementTop.md)
- [PickUp](Excel.Shape.PickUp.md)
- [RerouteConnections](Excel.Shape.RerouteConnections.md)
- [ScaleHeight](Excel.Shape.ScaleHeight.md)
- [ScaleWidth](Excel.Shape.ScaleWidth.md)
- [Select](Excel.Shape.Select.md)
- [SetShapesDefaultProperties](Excel.Shape.SetShapesDefaultProperties.md)
- [Ungroup](Excel.Shape.Ungroup.md)
- [ZOrder](Excel.Shape.ZOrder.md)

## Properties

- [Adjustments](Excel.Shape.Adjustments.md)
- [AlternativeText](Excel.Shape.AlternativeText.md)
- [Application](Excel.Shape.Application.md)
- [AutoShapeType](Excel.Shape.AutoShapeType.md)
- [BackgroundStyle](Excel.Shape.BackgroundStyle.md)
- [BlackWhiteMode](Excel.Shape.BlackWhiteMode.md)
- [BottomRightCell](Excel.Shape.BottomRightCell.md)
- [Callout](Excel.Shape.Callout.md)
- [Chart](Excel.Shape.Chart.md)
- [Child](Excel.Shape.Child.md)
- [ConnectionSiteCount](Excel.Shape.ConnectionSiteCount.md)
- [Connector](Excel.Shape.Connector.md)
- [ConnectorFormat](Excel.Shape.ConnectorFormat.md)
- [ControlFormat](Excel.Shape.ControlFormat.md)
- [Creator](Excel.Shape.Creator.md)
- [Decorative](Excel.Shape.Decorative.md)
- [Fill](Excel.Shape.Fill.md)
- [FormControlType](Excel.Shape.FormControlType.md)
- [Glow](Excel.Shape.Glow.md)
- [GraphicStyle](Excel.Shape.GraphicStyle.md)
- [GroupItems](Excel.Shape.GroupItems.md)
- [HasChart](Excel.Shape.HasChart.md)
- [HasSmartArt](Excel.Shape.HasSmartArt.md)
- [Height](Excel.Shape.Height.md)
- [HorizontalFlip](Excel.Shape.HorizontalFlip.md)
- [Hyperlink](Excel.Shape.Hyperlink.md)
- [ID](Excel.Shape.ID.md)
- [Left](Excel.Shape.Left.md)
- [Line](Excel.Shape.Line.md)
- [LinkFormat](Excel.Shape.LinkFormat.md)
- [LockAspectRatio](Excel.Shape.LockAspectRatio.md)
- [Locked](Excel.Shape.Locked.md)
- [Model3D](Excel.Shape.Model3D.md)
- [Name](Excel.Shape.Name.md)
- [Nodes](Excel.Shape.Nodes.md)
- [OLEFormat](Excel.Shape.OLEFormat.md)
- [OnAction](Excel.Shape.OnAction.md)
- [Parent](Excel.Shape.Parent.md)
- [ParentGroup](Excel.Shape.ParentGroup.md)
- [PictureFormat](Excel.Shape.PictureFormat.md)
- [Placement](Excel.Shape.Placement.md)
- [Reflection](Excel.Shape.Reflection.md)
- [Rotation](Excel.Shape.Rotation.md)
- [Shadow](Excel.Shape.Shadow.md)
- [ShapeStyle](Excel.Shape.ShapeStyle.md)
- [SmartArt](Excel.Shape.SmartArt.md)
- [SoftEdge](Excel.Shape.SoftEdge.md)
- [TextEffect](Excel.Shape.TextEffect.md)
- [TextFrame](Excel.Shape.TextFrame.md)
- [TextFrame2](Excel.Shape.TextFrame2.md)
- [ThreeD](Excel.Shape.ThreeD.md)
- [Title](Excel.Shape.Title.md)
- [Top](Excel.Shape.Top.md)
- [TopLeftCell](Excel.Shape.TopLeftCell.md)
- [Type](Excel.Shape.Type.md)
- [VerticalFlip](Excel.Shape.VerticalFlip.md)
- [Vertices](Excel.Shape.Vertices.md)
- [Visible](Excel.Shape.Visible.md)
- [Width](Excel.Shape.Width.md)
- [ZOrderPosition](Excel.Shape.ZOrderPosition.md)



## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
