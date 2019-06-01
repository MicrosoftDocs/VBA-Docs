---
title: Shape object (Publisher)
keywords: vbapb10.chm2293759
f1_keywords:
- vbapb10.chm2293759
ms.prod: publisher
api_name:
- Publisher.Shape
ms.assetid: 666cb7f0-62a8-f419-9838-007ef29506ee
ms.date: 06/01/2019
localization_priority: Normal
---


# Shape object (Publisher)

Represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, ActiveX control, or picture. The **Shape** object is a member of the **[Shapes](Publisher.Shapes.md)** collection, which includes all the shapes on a page or in a selection.

> [!NOTE] 
> There are three objects that represent shapes: 
> - The **Shapes** collection, which represents all the shapes on a document.
> - The **[ShapeRange](Publisher.ShapeRange.md)** collection, which represents a specified subset of the shapes on a document (for example, a **ShapeRange** object could represent shapes one and four on the document, or it could represent all the selected shapes on the document).
> - The **Shape** object, which represents a single shape on a document. 
> 
> If you want to work with several shapes at the same time or with shapes within the selection, use a **ShapeRange** collection. 

   
## Remarks

### Return an existing shape on a document

Use **[Shapes](Publisher.Shapes.md)** (_index_), where _index_ is the name or the index number, to return a single **Shape** object.

Each shape is assigned a default name when it is created. For example, if you add three different shapes to a document, they might be named Rectangle 2, TextBox 3, and Oval 4. To give a shape a more meaningful name, set the **[Name](publisher.shape.name.md)** property of the shape.

### Return a shape or shapes within a selection

Use **[Selection.ShapeRange](Publisher.Selection.ShapeRange.md)** (_index_), where _index_ is the name or the index number, to return a **Shape** object that represents a shape within a selection. 

### Return a newly created shape

To add a **Shape** object to the collection of shapes for the specified document and return a **Shape** object that represents the newly created shape, use one of the following methods of the **Shapes** collection: 

- **[AddCallout](Publisher.Shapes.AddCallout.md)**
- **[AddConnector](Publisher.Shapes.AddConnector.md)**
- **[AddCurve](Publisher.Shapes.AddCurve.md)**
- **[AddLabel](Publisher.Shapes.AddLabel.md)**
- **[AddLine](Publisher.Shapes.AddLine.md)**
- **[AddOLEObject](Publisher.Shapes.AddOLEObject.md)**
- **[AddPolyline](Publisher.Shapes.AddPolyline.md)**
- **[AddShape](Publisher.Shapes.AddShape.md)**
- **[AddTextBox](Publisher.Shapes.AddTextbox.md)**
- **[AddTextEffect](Publisher.Shapes.AddTextEffect.md)**

### Work with a group of shapes

Use **[GroupItems](Publisher.Shape.GroupItems.md)** (_index_), where _index_ is the shape name or the index number within the group, to return a **Shape** object that represents a single shape in a grouped shape. Use the **[ShapeRange.Group](Publisher.ShapeRange.Group.md)** or **[Regroup](Publisher.ShapeRange.Regroup.md)** method to group a range of shapes and return a single **Shape** object that represents the newly formed group. After a group has been formed, you can work with the group the same way that you work with any other shape. 

### Format a shape

- Use the **[AutoShapeType](Publisher.Shape.AutoShapeType.md)** property to specify the type of AutoShape: oval, rectangle, or balloon, for example.

- Use the **[Callout](Publisher.Shape.Callout.md)** property, which returns the **[CalloutFormat](Publisher.CalloutFormat.md)** object, to format line callouts. 

- Use the **[Fill](Publisher.Shape.Fill.md)** property to return the **[FillFormat](Publisher.FillFormat.md)** object, which contains all the properties and methods for formatting the fill of a closed shape. 

- Use the **[Line](Publisher.Shape.Line.md)** property to return a **[LineFormat](Publisher.LineFormat.md)** object, which contains properties and methods for formatting lines and arrows. 

- Use the **[PickUp](Publisher.Shape.PickUp.md)** and **[Apply](Publisher.Shape.Apply.md)** methods to transfer formatting from one shape to another.

- Use the **[SetShapesDefaultProperties](Publisher.Shape.SetShapesDefaultProperties.md)** method to set the formatting for the default shape for the document. New shapes inherit many of their attributes from the default shape.

- Use the **[Shadow](Publisher.Shape.Shadow.md)** property, which returns the **[ShadowFormat](Publisher.ShadowFormat.md)** object, to format a shadow. 

- Use the **[TextEffect](Publisher.Shape.TextEffect.md)** property, which returns the **[TextEffectFormat](Publisher.TextEffectFormat.md)** object, to format WordArt. 

- Use **[TextFrame](Publisher.Shape.TextFrame.md)** and **[Cell.TextRange](Publisher.Cell.TextRange.md)** properties to return the **[TextFrame](Publisher.TextFrame.md)** and **[TextRange](Publisher.TextRange.md)** objects, respectively, which contain all the properties and methods for inserting and formatting text within shapes and publications and linking the text frames together. 

- Use the **[TextWrap](Publisher.Shape.TextWrap.md)** property, which returns the **[WrapFormat](Publisher.WrapFormat.md)** object, to define how text wraps around shapes. 

- Use the **[ThreeD](Publisher.Shape.ThreeD.md)** property, which returns the **[ThreeDFormat](Publisher.ThreeDFormat.md)** object, to create 3D shapes. 

- Use the **[Type](Publisher.Shape.Type.md)** property to specify the type of shape: freeform, AutoShape, OLE object, callout, or linked picture, for example. 

- Use the **[Width](Publisher.Shape.Width.md)** and **[Height](Publisher.Shape.Height.md)** properties to specify the size of the shape.



## Example

The following example horizontally flips shape one on the active document.

```vb
Sub FlipShape() 
    ActiveDocument.Pages(1).Shapes(1).Flip FlipCmd:=msoFlipHorizontal 
End Sub
```

<br/>

The following example horizontally flips the shape named Rectangle 1 on the active document.

```vb
Sub FlipShapeByName() 
    ActiveDocument.Pages(1).Shapes("Rectangle 1") _ 
        .Flip FlipCmd:=msoFlipHorizontal 
End Sub
```

<br/>

The following example sets the fill for the first shape in the selection, assuming that the selection contains at least one shape.

```vb
Sub FillSelectedShape() 
    Selection.ShapeRange(1).Fill.ForeColor.RGB = RGB(255, 0, 0) 
End Sub
```

<br/>

The following example sets the fill for all the shapes in the selection, assuming that the selection contains at least one shape.

```vb
Sub FillAllSelectedShapes() 
    Dim shpShape As Shape 
    For Each
```


```vb
shpShape In Selection.ShapeRange 
       
```


```vb
shpShape.Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
    Next shpShape 
End Sub
```

<br/>

The following example adds a rectangle to the active document.

```vb
Sub AddNewShape() 
    ActiveDocument.Pages(1).Shapes.AddShape Type:=msoShapeRectangle, _ 
        Left:=400, Top:=72, Width:=100, Height:=200 
End Sub
```

<br/>

This example adds three shapes to the active publication, groups the shapes, and sets the fill color for each of the shapes in the group.

```vb
Sub WorkWithGroupShapes() 
 
    With ActiveDocument.Pages(1).Shapes 
        .AddShape Type:=msoShapeIsoscelesTriangle, Left:=100, _ 
            Top:=72, Width:=100, Height:=100 
        .AddShape Type:=msoShapeIsoscelesTriangle, Left:=250, _ 
            Top:=72, Width:=100, Height:=100 
        .AddShape Type:=msoShapeIsoscelesTriangle, Left:=400, _ 
            Top:=72, Width:=100, Height:=100 
        .SelectAll 
 
        With Selection.ShapeRange 
            .Group 
            .GroupItems(1).Fill.ForeColor _ 
                .RGB = RGB(Red:=255, Green:=0, Blue:=0) 
            .GroupItems(2).Fill.ForeColor _ 
                .RGB = RGB(Red:=0, Green:=255, Blue:=0) 
            .GroupItems(3).Fill.ForeColor _ 
                .RGB = RGB(Red:=0, Green:=0, Blue:=255) 
        End With 
    End With 
End Sub
```

<br/>

The following example adds a text box to the first page of the active publication, and then adds text to it and formats the text.

```vb
Sub CreateNewTextBox() 
    With ActiveDocument.Pages(1).Shapes.AddTextbox( _ 
        Orientation:=pbTextOrientationHorizontal, Left:=100, _ 
        Top:=100, Width:=200, Height:=100).TextFrame.TextRange 
        .Text = "This is a textbox." 
        With .Font 
            .Name = "Stencil" 
            .Bold = msoTrue 
            .Size = 30 
        End With 
    End With 
End Sub
```


## Methods

- [AddToCatalogMergeArea](Publisher.Shape.AddToCatalogMergeArea.md)
- [Apply](Publisher.Shape.Apply.md)
- [Copy](Publisher.Shape.Copy.md)
- [Cut](Publisher.Shape.Cut.md)
- [Delete](Publisher.Shape.Delete.md)
- [Duplicate](Publisher.Shape.Duplicate.md)
- [Flip](Publisher.Shape.Flip.md)
- [GetHeight](Publisher.Shape.GetHeight.md)
- [GetLeft](Publisher.Shape.GetLeft.md)
- [GetTop](Publisher.Shape.GetTop.md)
- [GetWidth](Publisher.Shape.GetWidth.md)
- [IncrementLeft](Publisher.Shape.IncrementLeft.md)
- [IncrementRotation](Publisher.Shape.IncrementRotation.md)
- [IncrementTop](Publisher.Shape.IncrementTop.md)
- [MoveIntoTextFlow](Publisher.Shape.MoveIntoTextFlow.md)
- [MoveOutOfTextFlow](Publisher.Shape.MoveOutOfTextFlow.md)
- [MoveToPage](Publisher.Shape.MoveToPage.md)
- [PickUp](Publisher.Shape.PickUp.md)
- [RemoveCatalogMergeArea](Publisher.Shape.RemoveCatalogMergeArea.md)
- [RemoveFromCatalogMergeArea](Publisher.Shape.RemoveFromCatalogMergeArea.md)
- [RerouteConnections](Publisher.Shape.RerouteConnections.md)
- [SaveAsBuildingBlock](Publisher.shape.saveasbuildingblock.md)
- [SaveAsPicture](Publisher.Shape.SaveAsPicture.md)
- [ScaleHeight](Publisher.Shape.ScaleHeight.md)
- [ScaleWidth](Publisher.Shape.ScaleWidth.md)
- [Select](Publisher.Shape.Select.md)
- [SetCaption](Publisher.shape.setcaption.md)
- [SetShapesDefaultProperties](Publisher.Shape.SetShapesDefaultProperties.md)
- [Ungroup](Publisher.Shape.Ungroup.md)
- [ZOrder](Publisher.Shape.ZOrder.md)

## Properties

- [Adjustments](Publisher.Shape.Adjustments.md)
- [AlternativeText](Publisher.Shape.AlternativeText.md)
- [Application](Publisher.Shape.Application.md)
- [AutoShapeType](Publisher.Shape.AutoShapeType.md)
- [BlackWhiteMode](Publisher.Shape.BlackWhiteMode.md)
- [BorderArt](Publisher.Shape.BorderArt.md)
- [Callout](Publisher.Shape.Callout.md)
- [CatalogMergeItems](Publisher.Shape.CatalogMergeItems.md)
- [ConnectionSiteCount](Publisher.Shape.ConnectionSiteCount.md)
- [Connector](Publisher.Shape.Connector.md)
- [ConnectorFormat](Publisher.Shape.ConnectorFormat.md)
- [Fill](Publisher.Shape.Fill.md)
- [Glow](Publisher.shape.glow.md)
- [GroupItems](Publisher.Shape.GroupItems.md)
- [HasTable](Publisher.Shape.HasTable.md)
- [HasTextFrame](Publisher.Shape.HasTextFrame.md)
- [Height](Publisher.Shape.Height.md)
- [HorizontalFlip](Publisher.Shape.HorizontalFlip.md)
- [Hyperlink](Publisher.Shape.Hyperlink.md)
- [ID](Publisher.Shape.ID.md)
- [InlineAlignment](Publisher.Shape.InlineAlignment.md)
- [InlineTextRange](Publisher.Shape.InlineTextRange.md)
- [IsExcess](Publisher.Shape.IsExcess.md)
- [IsGroupMember](Publisher.Shape.IsGroupMember.md)
- [IsInline](Publisher.Shape.IsInline.md)
- [Left](Publisher.Shape.Left.md)
- [Line](Publisher.Shape.Line.md)
- [LinkFormat](Publisher.Shape.LinkFormat.md)
- [LockAspectRatio](Publisher.Shape.LockAspectRatio.md)
- [Name](Publisher.Shape.Name.md)
- [Nodes](Publisher.Shape.Nodes.md)
- [OLEFormat](Publisher.Shape.OLEFormat.md)
- [Parent](Publisher.Shape.Parent.md)
- [ParentGroupShape](Publisher.Shape.ParentGroupShape.md)
- [PictureFormat](Publisher.Shape.PictureFormat.md)
- [Reflection](Publisher.shape.reflection.md)
- [Rotation](Publisher.Shape.Rotation.md)
- [Shadow](Publisher.Shape.Shadow.md)
- [SoftEdge](Publisher.shape.softedge.md)
- [Table](Publisher.Shape.Table.md)
- [Tags](Publisher.Shape.Tags.md)
- [TextEffect](Publisher.Shape.TextEffect.md)
- [TextFrame](Publisher.Shape.TextFrame.md)
- [TextWrap](Publisher.Shape.TextWrap.md)
- [ThreeD](Publisher.Shape.ThreeD.md)
- [Top](Publisher.Shape.Top.md)
- [Type](Publisher.Shape.Type.md)
- [VerticalFlip](Publisher.Shape.VerticalFlip.md)
- [Vertices](Publisher.Shape.Vertices.md)
- [WebCheckBox](Publisher.Shape.WebCheckBox.md)
- [WebCommandButton](Publisher.Shape.WebCommandButton.md)
- [WebListBox](Publisher.Shape.WebListBox.md)
- [WebNavigationBarSetName](Publisher.Shape.WebNavigationBarSetName.md)
- [WebOptionButton](Publisher.Shape.WebOptionButton.md)
- [WebTextBox](Publisher.Shape.WebTextBox.md)
- [Width](Publisher.Shape.Width.md)
- [Wizard](Publisher.Shape.Wizard.md)
- [WizardTag](Publisher.Shape.WizardTag.md)
- [WizardTagInstance](Publisher.Shape.WizardTagInstance.md)
- [ZOrderPosition](Publisher.Shape.ZOrderPosition.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]