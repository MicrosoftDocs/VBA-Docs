---
title: PictureFormat.FileSize property (Publisher)
keywords: vbapb10.chm3604757
f1_keywords:
- vbapb10.chm3604757
ms.prod: publisher
api_name:
- Publisher.PictureFormat.FileSize
ms.assetid: 8bad7bc0-7381-9bd8-3db8-5841e41ccb34
ms.date: 06/12/2019
localization_priority: Normal
---


# PictureFormat.FileSize property (Publisher)

Returns a **Long** that represents, in bytes, the size of the picture or OLE object as it appears in the specified publication. Read-only.


## Syntax

_expression_.**FileSize**

_expression_ A variable that represents a **[PictureFormat](Publisher.PictureFormat.md)** object.


## Return value

Long


## Remarks

If the picture or OLE object is linked, use the **[OriginalFileSize](Publisher.PictureFormat.OriginalFileSize.md)** property to determine the size of the linked file.

To determine whether a shape represents a linked picture, use either the **[Type](Publisher.Shape.Type.md)** property of the **Shape** object or the **[IsLinked](Publisher.PictureFormat.IsLinked.md)** property of the **PictureFormat** object.


## Example

The following example tests each picture in the active publication, and prints selected image properties for pictures that are linked.

```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 
 Debug.Print "File Name: " & .Filename 
 Debug.Print "Original File Size: " & .OriginalFileSize & " bytes" 
 Debug.Print "File size in publication: " & .FileSize & " bytes" 
 End With 
 End If 
 Next shpLoop 
Next pgLoop 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]