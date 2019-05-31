---
title: Shapes object (Publisher)
keywords: vbapb10.chm2228223
f1_keywords:
- vbapb10.chm2228223
ms.prod: publisher
api_name:
- Publisher.Shapes
ms.assetid: 52e069a6-d54b-a11a-1cba-96174329cb02
ms.date: 06/01/2019
localization_priority: Normal
---


# Shapes object (Publisher)

A collection of **[Shape](Publisher.Shape.md)** objects that represent all the shapes on a page of a publication. Each **Shape** object represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.

> [!NOTE] 
> If you want to work with a subset of the shapes on a document&mdash;for example, to do something to only the AutoShapes on the document or to only the selected shapes&mdash;you must construct a **[ShapeRange](Publisher.ShapeRange.md)** collection that contains the shapes with which you want to work.

## Remarks

Use the **[Page.Shapes](Publisher.Page.Shapes.md)** property to return the **Shapes** collection. Use **Shapes** (_index_), where _index_ is the index number, to return a single **Shape** object. 

> [!NOTE] 
> If you want to do something (such as delete or set a property) to all the shapes in a publication at the same time, use the **[Range](Publisher.Shapes.Range.md)** method to create a **ShapeRange** object that contains all the shapes in the **Shapes** collection, and then apply the appropriate property or method to the **ShapeRange** object.

Use one of the following methods of the **Shapes** collection to add a shape to a publication and return a **Shape** object that represents the newly created shape: 

- **[AddCallout](Publisher.Shapes.AddCallout.md)**
- **[AddConnector](Publisher.Shapes.AddConnector.md)**
- **[AddCurve](Publisher.Shapes.AddCurve.md)**
- **[AddLabel](Publisher.Shapes.AddLabel.md)**
- **[AddLine](Publisher.Shapes.AddLine.md)**
- **[AddOLEObject](Publisher.Shapes.AddOLEObject.md)**
- **[AddPolyline](Publisher.Shapes.AddPolyline.md)**
- **[AddShape](Publisher.Shapes.AddShape.md)**
- **[AddTextbox](Publisher.Shapes.AddTextbox.md)**
- **[AddTextEffect](Publisher.Shapes.AddTextEffect.md)** 


## Example

The following example selects all the shapes on the first page of the active publication.

```vb
Sub SelectAllShapes() 
    ActiveDocument.Pages(1).Shapes.SelectAll 
End Sub
```

<br/>

The following example adds a new shape to the active publication.

```vb
Sub AddNewShape() 
    ActiveDocument.Pages(1).Shapes.AddShape Type:=msoShapeFoldedCorner, _ 
        Left:=50, Top:=50, Width:=100, Height:=200 
End Sub
```

<br/>

The following example horizontally flips shape one on the first page of the active publication.

```vb
Sub FlipShape() 
    ActiveDocument.Pages(1).Shapes(1).Flip FlipCmd:=msoFlipHorizontal 
End Sub
```


## Methods

- [AddBuildingBlock](Publisher.shapes.addbuildingblock.md)
- [AddCallout](Publisher.Shapes.AddCallout.md)
- [AddCatalogMergeArea](Publisher.Shapes.AddCatalogMergeArea.md)
- [AddCatalogMergeFieldToCanvas](Publisher.Shapes.AddCatalogMergeFieldToCanvas.md)
- [AddConnector](Publisher.Shapes.AddConnector.md)
- [AddCurve](Publisher.Shapes.AddCurve.md)
- [AddEmptyPictureFrame](Publisher.Shapes.AddEmptyPictureFrame.md)
- [AddGroupWizard](Publisher.Shapes.AddGroupWizard.md)
- [AddLabel](Publisher.Shapes.AddLabel.md)
- [AddLine](Publisher.Shapes.AddLine.md)
- [AddOLEObject](Publisher.Shapes.AddOLEObject.md)
- [AddPicture](Publisher.Shapes.AddPicture.md)
- [AddPolyline](Publisher.Shapes.AddPolyline.md)
- [AddShape](Publisher.Shapes.AddShape.md)
- [AddTable](Publisher.Shapes.AddTable.md)
- [AddTextbox](Publisher.Shapes.AddTextbox.md)
- [AddTextEffect](Publisher.Shapes.AddTextEffect.md)
- [AddWebControl](Publisher.Shapes.AddWebControl.md)
- [AddWebNavigationBar](Publisher.Shapes.AddWebNavigationBar.md)
- [AddWordArt](Publisher.Shapes.AddWordArt.md)
- [BuildFreeform](Publisher.Shapes.BuildFreeform.md)
- [FindShapeByWizardTag](Publisher.Shapes.FindShapeByWizardTag.md)
- [Item](Publisher.Shapes.Item.md)
- [Paste](Publisher.Shapes.Paste.md)
- [Range](Publisher.Shapes.Range.md)
- [SelectAll](Publisher.Shapes.SelectAll.md)

## Properties

- [Application](Publisher.Shapes.Application.md)
- [CanvasArrangementType](Publisher.Shapes.CanvasArrangementType.md)
- [CanvasesCount](Publisher.Shapes.CanvasesCount.md)
- [Count](Publisher.Shapes.Count.md)
- [Parent](Publisher.Shapes.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]