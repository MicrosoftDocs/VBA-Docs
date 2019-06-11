---
title: PictureFormat.HasTransparencyColor property (Publisher)
keywords: vbapb10.chm3604789
f1_keywords:
- vbapb10.chm3604789
ms.prod: publisher
api_name:
- Publisher.PictureFormat.HasTransparencyColor
ms.assetid: 2e6066e8-60b0-c33e-0bb0-1b6f83208fd0
ms.date: 06/12/2019
localization_priority: Normal
---


# PictureFormat.HasTransparencyColor property (Publisher)

Returns a **Boolean** that indicates whether a transparency color has been applied to the specified picture. Read-only.


## Syntax

_expression_.**HasTransparencyColor**

_expression_ A variable that represents a **[PictureFormat](Publisher.PictureFormat.md)** object.


## Return value

Boolean


## Example

The following example returns a list of the pictures with transparency colors in the active publication.

```vb
Sub ListPicturesWithTransColors() 
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
 For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 
 If shpLoop.Type = pbPicture Or shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse Then 
 If .HasTransparencyColor = True Then 
 Debug.Print .Filename 
 End If 
 End If 
 End With 
 
 End If 
 
 Next shpLoop 
 Next pgLoop 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]