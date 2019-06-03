---
title: CatalogMergeShapes object (Publisher)
keywords: vbapb10.chm8454143
f1_keywords:
- vbapb10.chm8454143
ms.prod: publisher
api_name:
- Publisher.CatalogMergeShapes
ms.assetid: 1108e9a4-57ef-2b1a-0998-54b6fad838da
ms.date: 05/31/2019
localization_priority: Normal
---


# CatalogMergeShapes object (Publisher)

Represents the shapes contained in the catalog merge area of the specified publication.
 
## Remarks

The catalog merge area is automatically resized to accommodate objects that are larger than the merge area, or that are positioned outside the catalog merge area when they are added.
 
Shapes inside the catalog merge area are automatically resized or repositioned if the catalog merge area is decreased in size or moved.

The catalog merge area can contain picture and text data fields that you have inserted in addition to other design elements that you choose. 

Use the **[CatalogMergeItems](Publisher.Shape.CatalogMergeItems.md)** property of the **Shape** or **[ShapeRange](Publisher.ShapeRange.md)** objects to return the contents of the catalog merge area. 

Use the **[AddToCatalogMergeArea](Publisher.Shape.AddToCatalogMergeArea.md)** method of the **Shape** or **ShapeRange** objects to add shapes to a catalog merge area. 

Use **CatalogMergeItems** (_index_), where _index_ is the index number, to return a single catalog merge area shape. 

Use the **[RemoveFromCatalogMergeArea](Publisher.Shape.RemoveFromCatalogMergeArea.md)** method of the **Shape** or **ShapeRange** objects to remove shapes from a catalog merge area. Removed shapes are not deleted, but are instead placed on the publication page containing the catalog merge area. 

## Example

The following example tests whether the specified publication contains a catalog merge area. If it does, it returns a list of the shapes that it contains.
 
```vb
Sub ListCatalogMergeAreaContents() 
 
 Dim pgPage As Page 
 Dim mmLoop As Shape 
 Dim intCount As Integer 
 
 For Each pgPage In ThisDocument.Pages 
 For Each mmLoop In pgPage.Shapes 
 
 If mmLoop.Type = pbCatalogMergeArea Then 
 
 With mmLoop.CatalogMergeItems 
 For intCount = 1 To .Count 
 Debug.Print "Shape ID: " & _ 
 mmLoop.CatalogMergeItems.Item(intCount).ID 
 Debug.Print "Shape Name: " & _ 
 mmLoop.CatalogMergeItems.Item(intCount).Name 
 Next 
 End With 
 
 End If 
 
 Next mmLoop 
 Next pgPage 
 
End Sub 

```

<br/>

The following example adds a rectangle to the catalog merge area in the specified publication. This example assumes that a catalog merge area has been added to the first page of the publication.
 
```vb
ThisDocument.Pages(1).Shapes.AddShape(1, 80, 75, 450, 125).AddToCatalogMergeArea
```

<br/>

The following example removes the first shape from the catalog merge area.

```vb
ThisDocument.Pages(1).Shapes(1).CatalogMergeItems(1).RemoveFromCatalogMergeArea
```

<br/>

The following example tests whether the specified publication contains a catalog merge area. If it does, all the shapes are removed from the catalog merge area and deleted, and the catalog merge area is then removed from the publication.

```vb
Sub DeleteCatalogMergeAreaAndAllShapesWithin() 
 Dim pgPage As Page 
 Dim mmLoop As Shape 
 Dim intCount As Integer 
 Dim strName As String 
 
 For Each pgPage In ThisDocument.Pages 
 For Each mmLoop In pgPage.Shapes 
 
 If mmLoop.Type = pbCatalogMergeArea Then 
 With mmLoop.CatalogMergeItems 
 For intCount = .Count To 1 Step -1 
 strName = mmLoop.CatalogMergeItems.Item(intCount).Name 
 .Item(intCount).RemoveFromCatalogMergeArea 
 pgPage.Shapes(strName).Delete 
 Next 
 End With 
 mmLoop.RemoveCatalogMergeArea 
 End If 
 
 Next mmLoop 
 Next pgPage 
 
 End Sub 

```


## Methods

- [Item](Publisher.CatalogMergeShapes.Item.md)
- [Range](Publisher.CatalogMergeShapes.Range.md)

## Properties

- [Application](Publisher.CatalogMergeShapes.Application.md)
- [Count](Publisher.CatalogMergeShapes.Count.md)
- [HorizontalRepeat](Publisher.CatalogMergeShapes.HorizontalRepeat.md)
- [Parent](Publisher.CatalogMergeShapes.Parent.md)
- [VerticalRepeat](Publisher.CatalogMergeShapes.VerticalRepeat.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]