---
title: ShapeRange.RemoveFromCatalogMergeArea method (Publisher)
keywords: vbapb10.chm2294049
f1_keywords:
- vbapb10.chm2294049
ms.prod: publisher
api_name:
- Publisher.ShapeRange.RemoveFromCatalogMergeArea
ms.assetid: 732cd277-9c2e-0a01-c2b5-8d016637884a
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.RemoveFromCatalogMergeArea method (Publisher)

Removes a shape from the specified page's catalog merge area. Removed shapes are not deleted, but instead remain in place on the page containing the catalog merge area.


## Syntax

_expression_.**RemoveFromCatalogMergeArea**

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Remarks

Use the **[AddToCatalogMergeArea](Publisher.ShapeRange.AddToCatalogMergeArea.md)** method of the **[Shape](Publisher.Shape.md)** or **ShapeRange** objects to add shapes to a catalog merge area.

Use the **[RemoveCatalogMergeArea](Publisher.Shape.RemoveCatalogMergeArea.md)** method of the **Shape** object to remove the catalog merge area from a publication page, but leave the shapes that it contains.


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