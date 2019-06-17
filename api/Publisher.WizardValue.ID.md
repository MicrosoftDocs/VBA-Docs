---
title: WizardValue.ID property (Publisher)
keywords: vbapb10.chm2097155
f1_keywords:
- vbapb10.chm2097155
ms.prod: publisher
api_name:
- Publisher.WizardValue.ID
ms.assetid: d8d1ec6b-e2e7-8729-b4d2-a62a578ead11
ms.date: 06/18/2019
localization_priority: Normal
---


# WizardValue.ID property (Publisher)

Returns a **Long** that represents the type of a shape, range of shapes, or property, type, or value of a wizard. Read-only.


## Syntax

_expression_.**ID**

_expression_ A variable that represents a **[WizardValue](Publisher.WizardValue.md)** object.


## Example

This example displays the type for each shape on the first page of the active publication.

```vb
Sub ShapeID() 
 Dim shp As Shape 
 For Each shp In ActiveDocument.Pages(1).Shapes 
 MsgBox shp.ID 
 Next shp 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]