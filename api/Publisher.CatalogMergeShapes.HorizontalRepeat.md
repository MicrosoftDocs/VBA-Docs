---
title: CatalogMergeShapes.HorizontalRepeat property (Publisher)
keywords: vbapb10.chm8388613
f1_keywords:
- vbapb10.chm8388613
ms.prod: publisher
api_name:
- Publisher.CatalogMergeShapes.HorizontalRepeat
ms.assetid: 1c3f1093-294f-e7b3-02ca-803ce7437d49
ms.date: 06/06/2019
localization_priority: Normal
---


# CatalogMergeShapes.HorizontalRepeat property (Publisher)

Returns a **Long** that represents the number of times that the catalog merge area repeats across the target publication page when the catalog merge is executed. Read-only.


## Syntax

_expression_.**HorizontalRepeat**

_expression_ A variable that represents a **[CatalogMergeShapes](Publisher.CatalogMergeShapes.md)** object.


## Return value

Long


## Remarks

When the catalog merge is executed, the catalog merge area repeats once for each selected record in the specified data source.

The number of times that the catalog merge area repeats across the page is determined by the width of the area. Use the **[Width](Publisher.Shape.Width.md)** property of the **Shape** object to return or set the horizontal size of the catalog merge area.

The **[VerticalRepeat](Publisher.CatalogMergeShapes.VerticalRepeat.md)** property represents the number of times that the catalog merge area repeats vertically down the target publication page.


## Example

The following example returns the number of times that the catalog merge area repeats horizontally and vertically on the target publication page when the catalog merge is performed. This example assumes that the catalog merge area is the first shape on the first page of the specified publication.

```vb
Sub CatalogMergeDimensions() 
 
 With ThisDocument.Pages(1).Shapes(1) 
 Debug.Print .Width 
 Debug.Print .CatalogMergeItems.HorizontalRepeat 
 Debug.Print .Height 
 Debug.Print .CatalogMergeItems.VerticalRepeat 
 End With 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]