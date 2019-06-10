---
title: OLEFormat.ProgId property (Publisher)
keywords: vbapb10.chm4456452
f1_keywords:
- vbapb10.chm4456452
ms.prod: publisher
api_name:
- Publisher.OLEFormat.ProgId
ms.assetid: dae7e591-65d2-b956-e598-8746955c4182
ms.date: 06/11/2019
localization_priority: Normal
---


# OLEFormat.ProgId property (Publisher)

Returns a **String** that represents the programmatic identifier (ProgID) for the specified OLE object. Read-only.


## Syntax

_expression_.**ProgId**

_expression_ A variable that represents an **[OLEFormat](Publisher.OLEFormat.md)** object.


## Return value

String


## Example

This example loops through all the linked OLE object shapes on the first page of the active document and updates all linked Excel worksheets. This example assumes that there is at least one shape on the first page of the active publication.

```vb
Sub UpdateLinkedOLEObject() 
 Dim shp As Shape 
 For Each shp In ActiveDocument.Pages(1).Shapes 
 If shp.Type = msoLinkedOLEObject Then 
 If shp.OLEFormat.ProgId = "Excel.Sheet" Then 
 shp.LinkFormat.Update 
 End If 
 End If 
 Next 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]