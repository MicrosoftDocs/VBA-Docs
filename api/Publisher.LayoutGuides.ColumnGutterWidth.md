---
title: LayoutGuides.ColumnGutterWidth property (Publisher)
keywords: vbapb10.chm1114128
f1_keywords:
- vbapb10.chm1114128
ms.prod: publisher
api_name:
- Publisher.LayoutGuides.ColumnGutterWidth
ms.assetid: 1c8fd297-1164-da50-cee8-390263cce5b0
ms.date: 06/08/2019
localization_priority: Normal
---


# LayoutGuides.ColumnGutterWidth property (Publisher)

Returns or sets the width of the column gutters that are used by the **LayoutGuides** object to aid in the process of laying out design elements. Read/write **Single**.


## Syntax

_expression_.**ColumnGutterWidth**

_expression_ A variable that represents a **[LayoutGuides](Publisher.LayoutGuides.md)** object.


## Return value

Single


## Remarks

The default width of column gutters is 0.4 inches.


## Example

The following example modifies the second master page of the active publication so that it has four rows and four columns, a row gutter width of 0.75 inches, a column gutter width of 0.5 inches, and center lines in the gutters. Any new pages added to the publication that use the second master page as a template have these properties.

```vb
Dim theMasterPage As page 
Dim theLayoutGuides As LayoutGuides 
 
Set theMasterPage = ActiveDocument.MasterPages(2) 
Set theLayoutGuides = theMasterPage.LayoutGuides 
 
With theLayoutGuides 
 .Rows = 4 
 .Columns = 4 
 .RowGutterWidth = Application.InchesToPoints(0.75) 
 .ColumnGutterWidth = Application.InchesToPoints(0.5) 
 .GutterCenterlines = True 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]