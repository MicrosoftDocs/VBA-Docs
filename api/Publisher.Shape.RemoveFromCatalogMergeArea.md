---
title: Shape.RemoveFromCatalogMergeArea method (Publisher)
keywords: vbapb10.chm5308689
f1_keywords:
- vbapb10.chm5308689
ms.prod: publisher
api_name:
- Publisher.Shape.RemoveFromCatalogMergeArea
ms.assetid: 3b3630c3-6bf1-494b-151c-c930f32a2a77
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.RemoveFromCatalogMergeArea method (Publisher)

Removes a shape from the specified page's catalog merge area. Removed shapes are not deleted, but instead remain in place on the page that contains the catalog merge area.


## Syntax

_expression_.**RemoveFromCatalogMergeArea**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Return value

Nothing


## Remarks

Use the **[AddToCatalogMergeArea](Publisher.Shape.AddToCatalogMergeArea.md)** method of the **Shape** or **[ShapeRange](Publisher.ShapeRange.md)** objects to add shapes to a catalog merge area.

Use the **[RemoveCatalogMergeArea](Publisher.Shape.RemoveCatalogMergeArea.md)** method to remove the catalog merge area from a publication page, but leave the shapes it contains.


## Example

The following example tests whether any page of the specified publication contains a catalog merge area. If any page does, all the shapes are removed from the catalog merge area and deleted, and the catalog merge area is then removed from the publication.

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

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]