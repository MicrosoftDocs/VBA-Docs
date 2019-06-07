---
title: LinkFormat.SourceFullName property (Publisher)
keywords: vbapb10.chm4390915
f1_keywords:
- vbapb10.chm4390915
ms.prod: publisher
api_name:
- Publisher.LinkFormat.SourceFullName
ms.assetid: a83aad48-ce27-6fe7-d26b-f00bec42e614
ms.date: 06/08/2019
localization_priority: Normal
---


# LinkFormat.SourceFullName property (Publisher)

Returns a **String** that represents the path and name of the source file for the specified linked OLE object, picture, or field. Read-only.


## Syntax

_expression_.**SourceFullName**

_expression_ A variable that represents a **[LinkFormat](Publisher.LinkFormat.md)** object.


## Return value

String


## Example

This example displays the path and file name of the source file for all embedded OLE shapes on the first page of the active publication.

```vb
Sub DisplaySourceName() 
 Dim shp As Shape 
 For Each shp In ActiveDocument.Pages(1).Shapes 
 If shp.Type = pbEmbeddedOLEObject Then 
 With shp.LinkFormat 
 MsgBox .SourceFullName 
 End With 
 End If 
 Next 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]