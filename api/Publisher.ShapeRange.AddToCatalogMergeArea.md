---
title: ShapeRange.AddToCatalogMergeArea method (Publisher)
keywords: vbapb10.chm2294048
f1_keywords:
- vbapb10.chm2294048
ms.prod: publisher
api_name:
- Publisher.ShapeRange.AddToCatalogMergeArea
ms.assetid: 6cb770c6-fe6e-ffe8-cd51-855d97b17aed
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.AddToCatalogMergeArea method (Publisher)

Adds the specified shape or shapes to the publication page's catalog merge area.


## Syntax

_expression_.**AddToCatalogMergeArea**

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Remarks

The catalog merge area is automatically resized to accommodate objects that are larger than the merge area, or that are positioned outside the catalog merge area when they are added.

The **AddToCatalogMergeArea** method does not apply to merge data fields:

- Use the **[Insert](Publisher.MailMergeDataField.Insert.md)** method of the **MailMergeDataField** object to add a picture data field to a publication page's catalog merge area.   
- Use the **[InsertMailMergeField](Publisher.TextRange.InsertMailMergeField.md)** method of the **TextRange** object to add a text data field to a text box.
    
Use the **AddToCatalogMergeArea** method to add a text box that contains text data fields to a catalog merge area. 


## Example

The following example adds a rectangle to the catalog merge area on the first page of the specified publication. This example assumes that a catalog merge area has been added to the first page.

```vb
ThisDocument.Pages(1).Shapes.AddShape(1, 80, 75, 450, 125).AddToCatalogMergeArea
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]