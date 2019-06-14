---
title: Shapes.AddCatalogMergeArea method (Publisher)
keywords: vbapb10.chm2162752
f1_keywords:
- vbapb10.chm2162752
ms.prod: publisher
api_name:
- Publisher.Shapes.AddCatalogMergeArea
ms.assetid: 4af86b99-5a3a-b9f3-d269-16d635d35c83
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.AddCatalogMergeArea method (Publisher)

Adds a **[Shape](Publisher.Shape.md)** object that represents the specified publication's catalog merge area.


## Syntax

_expression_.**AddCatalogMergeArea**

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


## Return value

Shape


## Remarks

Only one catalog merge area can be added to a publication page. Typically, a publication will only have one catalog merge area.

Although you can add one catalog merge area per publication page, you can only connect to a single data source for a publication. What data is merged is determined by the catalog merge area on the active page, and the data fields it contains.

> [!NOTE] 
> You must add a catalog merge area to the publication page before you connect to a data source.

Use the **[AddToCatalogMergeArea](Publisher.Shape.AddToCatalogMergeArea.md)** method of the **Shape** or **[ShapeRange](Publisher.ShapeRange.md)** objects to add shapes to a catalog merge area.

Use the **[Insert](Publisher.MailMergeDataField.Insert.md)** method of the **MailMergeDataField** object to add a picture data field to a publication's catalog merge area.

Use the **[InsertMailMergeField](Publisher.TextRange.InsertMailMergeField.md)** method of the **TextRange** object to add a text data field to a text box in the publication's catalog merge area.

Use the **[RemoveCatalogMergeArea](Publisher.Shape.RemoveCatalogMergeArea.md)** method of the **Shape** object to remove a catalog merge area from a publication.

This method corresponds to selecting a catalog merge in **Step 1: Select a merge type** of the Mail and Catalog Merge Wizard.


## Example

The following example adds a catalog merge area to the first page of the specified publication.

```vb
ThisDocument.Pages(1).Shapes.AddCatalogMergeArea
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]