---
title: ShapeRange object (Publisher)
keywords: vbapb10.chm2359295
f1_keywords:
- vbapb10.chm2359295
ms.prod: publisher
api_name:
- Publisher.ShapeRange
ms.assetid: c85967c9-af43-747d-7e0b-64ddc22c84be
ms.date: 06/01/2019
localization_priority: Normal
---


# ShapeRange object (Publisher)

Represents a shape range, which is a set of shapes on a document. A shape range can contain as few as one shape or as many as all the shapes in the document. You can include whichever shapes you want&mdash;chosen from among all the shapes in the document or all the shapes in the selection&mdash;to construct a shape range. For example, you could construct a **ShapeRange** collection that contains the first three shapes in a document, all the selected shapes in a document, or all the freeform shapes in a document.

> [!NOTE] 
> Most operations that you can do with a **[Shape](Publisher.Shape.md)** object, you can also do with a **ShapeRange** object that contains only one shape. Some operations, when performed on a **ShapeRange** object that contains more than one shape, cause an error. 
    
## Remarks

Use **[Shapes.Range](Publisher.Shapes.Range.md)** (_index_), where _index_ is the index number of the shape or an array that contains index numbers of shapes, to return a **ShapeRange** collection that represents a set of shapes in a publication. You can use Visual Basic's **Array** function to construct an array of index numbers. 

Although you can use the **Shapes.Range** method to return any number of shapes, it is simpler to use the **[Item](Publisher.ShapeRange.Item.md)** method if you want to return only a single member of the collection. For example, **Shapes** (1) is simpler than **Shapes.Range** (1).

Use **[Selection.ShapeRange](Publisher.Selection.ShapeRange.md)** (_index_), where _index_ is the index number of the shape, to return a **Shape** object that represents a shape within a selection. 

Use the **[Align](Publisher.ShapeRange.Align.md)** method, **[Distribute](Publisher.ShapeRange.Distribute.md)** method, or **[ZOrder](Publisher.ShapeRange.ZOrder.md)** method to position a set of shapes relative to each other or relative to the document. 

Use the **[Group](Publisher.ShapeRange.Group.md)** method, **[Regroup](Publisher.ShapeRange.Regroup.md)** method, or **[Ungroup](Publisher.ShapeRange.Ungroup.md)** method to create and work with a single shape formed from a shape range. The **[GroupItems](Publisher.ShapeRange.GroupItems.md)** property returns the **[GroupShapes](Publisher.GroupShapes.md)** object, which represents all the shapes that were grouped to form one shape. 


## Example

The following example sets the fill pattern for shapes one through three on the active publication.

```vb
Sub ChangeFillPattern() 
    ActiveDocument.Pages(1).Shapes.Range(Array(1, 2, 3)) _ 
        .Fill.PresetGradient Style:=msoGradientDiagonalDown, _ 
        Variant:=1, PresetGradientType:=msoGradientHorizon 
End Sub
```

<br/>

The following example selects the first two shapes on the first page of the active publication and then sets the fill for the first shape in the selection.

```vb
Sub ChangeFillForShapeRange() 
    ActiveDocument.Pages(1).Shapes.Range(Array(1, 2)).Select 
    Selection.ShapeRange(1).Fill.ForeColor.RGB = RGB(255, 0, 0) 
End Sub
```

<br/>

This example selects all the shapes on the first page of the active publication, and then adds and formats text in the second shape in the range.

```vb
Sub SelectShapesOnPageOne() 
    ActiveDocument.Pages(1).Shapes.Range.Select 
    With Selection.ShapeRange(2).TextFrame.TextRange 
        .Text = "Shape Number 2" 
        .ParagraphFormat.Alignment = pbParagraphAlignmentCenter 
        .Font.Size = 25 
    End With 
End Sub
```

<br/>

This example specifies a shape range and left-aligns and vertically distributes the shapes on the page.

```vb
Sub AlignDistributeShapes() 
    Dim rngShapes As ShapeRange 
    Set rngShapes = ActiveDocument.Pages(1).Shapes.Range 
 
    With rngShapes 
        .Align AlignCmd:=msoAlignLefts, RelativeTo:=msoFalse 
        .Distribute DistributeCmd:=msoDistributeVertically, RelativeTo:=msoTrue 
    End With 
End Sub
```

<br/>

This example specifies a shape range and left-aligns and vertically distributes the shapes on the page.

```vb
Sub GroupShapes() 
    Dim rngShapes As ShapeRange 
    Set rngShapes = ActiveDocument.Pages(1).Shapes.Range 
    rngShapes.Group 
 
    rngShapes(1).Fill.OneColorGradient _ 
        Style:=msoGradientFromCenter, _ 
        Variant:=2, Degree:=1 
End Sub
```


## Methods

- [AddToCatalogMergeArea](Publisher.ShapeRange.AddToCatalogMergeArea.md)
- [Align](Publisher.ShapeRange.Align.md)
- [Apply](Publisher.ShapeRange.Apply.md)
- [Copy](Publisher.ShapeRange.Copy.md)
- [Cut](Publisher.ShapeRange.Cut.md)
- [Delete](Publisher.ShapeRange.Delete.md)
- [Distribute](Publisher.ShapeRange.Distribute.md)
- [Duplicate](Publisher.ShapeRange.Duplicate.md)
- [Flip](Publisher.ShapeRange.Flip.md)
- [GetHeight](Publisher.ShapeRange.GetHeight.md)
- [GetLeft](Publisher.ShapeRange.GetLeft.md)
- [GetTop](Publisher.ShapeRange.GetTop.md)
- [GetWidth](Publisher.ShapeRange.GetWidth.md)
- [Group](Publisher.ShapeRange.Group.md)
- [IncrementLeft](Publisher.ShapeRange.IncrementLeft.md)
- [IncrementRotation](Publisher.ShapeRange.IncrementRotation.md)
- [IncrementTop](Publisher.ShapeRange.IncrementTop.md)
- [Item](Publisher.ShapeRange.Item.md)
- [MoveIntoTextFlow](Publisher.ShapeRange.MoveIntoTextFlow.md)
- [MoveOutOfTextFlow](Publisher.ShapeRange.MoveOutOfTextFlow.md)
- [PickUp](Publisher.ShapeRange.PickUp.md)
- [Regroup](Publisher.ShapeRange.Regroup.md)
- [RemoveFromCatalogMergeArea](Publisher.ShapeRange.RemoveFromCatalogMergeArea.md)
- [RerouteConnections](Publisher.ShapeRange.RerouteConnections.md)
- [SaveAsBuildingBlock](Publisher.shaperange.saveasbuildingblock.md)
- [SaveAsPicture](Publisher.ShapeRange.SaveAsPicture.md)
- [ScaleHeight](Publisher.ShapeRange.ScaleHeight.md)
- [ScaleWidth](Publisher.ShapeRange.ScaleWidth.md)
- [Select](Publisher.ShapeRange.Select.md)
- [SetShapesDefaultProperties](Publisher.ShapeRange.SetShapesDefaultProperties.md)
- [Ungroup](Publisher.ShapeRange.Ungroup.md)
- [ZOrder](Publisher.ShapeRange.ZOrder.md)

## Properties

- [Adjustments](Publisher.ShapeRange.Adjustments.md)
- [AlternativeText](Publisher.ShapeRange.AlternativeText.md)
- [Application](Publisher.ShapeRange.Application.md)
- [AutoShapeType](Publisher.ShapeRange.AutoShapeType.md)
- [BlackWhiteMode](Publisher.ShapeRange.BlackWhiteMode.md)
- [Callout](Publisher.ShapeRange.Callout.md)
- [ConnectionSiteCount](Publisher.ShapeRange.ConnectionSiteCount.md)
- [Connector](Publisher.ShapeRange.Connector.md)
- [ConnectorFormat](Publisher.ShapeRange.ConnectorFormat.md)
- [Count](Publisher.ShapeRange.Count.md)
- [Fill](Publisher.ShapeRange.Fill.md)
- [Glow](Publisher.shaperange.glow.md)
- [GroupItems](Publisher.ShapeRange.GroupItems.md)
- [HasTable](Publisher.ShapeRange.HasTable.md)
- [HasTextFrame](Publisher.ShapeRange.HasTextFrame.md)
- [Height](Publisher.ShapeRange.Height.md)
- [HorizontalFlip](Publisher.ShapeRange.HorizontalFlip.md)
- [Hyperlink](Publisher.ShapeRange.Hyperlink.md)
- [ID](Publisher.ShapeRange.ID.md)
- [InlineAlignment](Publisher.ShapeRange.InlineAlignment.md)
- [InlineTextRange](Publisher.ShapeRange.InlineTextRange.md)
- [IsInline](Publisher.ShapeRange.IsInline.md)
- [Left](Publisher.ShapeRange.Left.md)
- [Line](Publisher.ShapeRange.Line.md)
- [LinkFormat](Publisher.ShapeRange.LinkFormat.md)
- [LockAspectRatio](Publisher.ShapeRange.LockAspectRatio.md)
- [Name](Publisher.ShapeRange.Name.md)
- [Nodes](Publisher.ShapeRange.Nodes.md)
- [OLEFormat](Publisher.ShapeRange.OLEFormat.md)
- [Parent](Publisher.ShapeRange.Parent.md)
- [PictureFormat](Publisher.ShapeRange.PictureFormat.md)
- [Reflection](Publisher.shaperange.reflection.md)
- [Rotation](Publisher.ShapeRange.Rotation.md)
- [Shadow](Publisher.ShapeRange.Shadow.md)
- [SoftEdge](Publisher.shaperange.softedge.md)
- [Table](Publisher.ShapeRange.Table.md)
- [Tags](Publisher.ShapeRange.Tags.md)
- [TextEffect](Publisher.ShapeRange.TextEffect.md)
- [TextFrame](Publisher.ShapeRange.TextFrame.md)
- [TextWrap](Publisher.ShapeRange.TextWrap.md)
- [ThreeD](Publisher.ShapeRange.ThreeD.md)
- [Top](Publisher.ShapeRange.Top.md)
- [Type](Publisher.ShapeRange.Type.md)
- [VerticalFlip](Publisher.ShapeRange.VerticalFlip.md)
- [Vertices](Publisher.ShapeRange.Vertices.md)
- [Width](Publisher.ShapeRange.Width.md)
- [Wizard](Publisher.ShapeRange.Wizard.md)
- [WizardTag](Publisher.ShapeRange.WizardTag.md)
- [WizardTagInstance](Publisher.ShapeRange.WizardTagInstance.md)
- [ZOrderPosition](Publisher.ShapeRange.ZOrderPosition.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]