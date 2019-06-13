---
title: PictureFormat.OriginalFileSize property (Publisher)
keywords: vbapb10.chm3604772
f1_keywords:
- vbapb10.chm3604772
ms.prod: publisher
api_name:
- Publisher.PictureFormat.OriginalFileSize
ms.assetid: 30704f2a-d739-7f14-d69a-73ab1f5ab8f3
ms.date: 06/13/2019
localization_priority: Normal
---


# PictureFormat.OriginalFileSize property (Publisher)

Returns a **Long** representing the size, in bytes, of the linked picture or OLE object. Read-only.


## Syntax

_expression_.**OriginalFileSize**

_expression_ A variable that represents a **[PictureFormat](Publisher.PictureFormat.md)** object.


## Remarks

This property only applies to linked pictures. Returns "Permission Denied" for shapes representing embedded or pasted pictures.

Use either of the following properties to determine whether a shape represents a linked picture:

- The **[Type](Publisher.Shape.Type.md)** property of the **Shape** object   
- The **[IsLinked](Publisher.PictureFormat.IsLinked.md)** property of the **PictureFormat** object
    
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
 
 End With 
 End If 
 Next shpLoop 
Next pgLoop 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]