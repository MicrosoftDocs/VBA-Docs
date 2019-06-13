---
title: PictureFormat.FileName property (Publisher)
keywords: vbapb10.chm3604756
f1_keywords:
- vbapb10.chm3604756
ms.prod: publisher
api_name:
- Publisher.PictureFormat.FileName
ms.assetid: 73e2a224-f15a-50cc-462e-10ccf9478122
ms.date: 06/12/2019
localization_priority: Normal
---


# PictureFormat.FileName property (Publisher)

Returns a **String** that represents the file name of the specified picture or OLE object. Read-only.


## Syntax

_expression_.**FileName**

_expression_ A variable that represents a **[PictureFormat](Publisher.PictureFormat.md)** object.


## Return value

String


## Remarks

For linked pictures and OLE objects, the returned string represents the full path and file name of the picture. For embedded pictures and OLE objects, the returned string represents the file name only.

To determine whether a shape represents a linked picture, use either the **[Type](Publisher.Shape.Type.md)** property of the **Shape** object or the **[IsLinked](Publisher.PictureFormat.IsLinked.md)** property of the **PictureFormat** object.


## Example

The following example returns selected image properties for each picture in the active publication.

```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbPicture Or shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 
 If .IsEmpty = msoFalse Then 
 
 Debug.Print "File Name: " & .FileName 
 Debug.Print "Horizontal Scaling: " & .HorizontalScale & "%" 
 Debug.Print "Vertical Scaling: " & .VerticalScale & "%" 
 Debug.Print "File size in publication: " & .FileSize & " bytes" 
 
 End If 
 
 End With 
 End If 
 Next shpLoop 
Next pgLoop
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]