---
title: PictureFormat.OriginalResolution Property (Publisher)
keywords: vbapb10.chm3604776
f1_keywords:
- vbapb10.chm3604776
ms.prod: publisher
api_name:
- Publisher.PictureFormat.OriginalResolution
ms.assetid: 0cb7ee4e-3eb8-baee-6535-d936e3c5f05c
ms.date: 06/08/2017
localization_priority: Normal
---


# PictureFormat.OriginalResolution Property (Publisher)

Returns a  **Long** that represents, in dots per inch (dpi), the resolution at which the linked picture was originally scanned. Read-only.


## Syntax

 _expression_. **OriginalResolution**

 _expression_ A variable that represents an  **PictureFormat** object.


## Return value

Long


## Remarks

This property only applies to linked pictures. Returns "Permission Denied" for shapes representing embedded or pasted pictures.

To determine whether a shape represents a linked picture, use either the  **[Type](Publisher.Shape.Type.md)** property of the **[Shape](Publisher.Shape.md)** object, or the **[IsLinked](Publisher.PictureFormat.IsLinked.md)** property of the **[PictureFormat](Publisher.PictureFormat.md)** object.

Use the  **[EffectiveResolution](Publisher.PictureFormat.EffectiveResolution.md)** property to determine the resolution at which the picture or OLE object prints in the specified document.


## Example

The following example tests each picture in the active publication, and returns selected image properties for pictures that are linked.


```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 
 Debug.Print "File Name: " & .Filename 
 Debug.Print "Resolution in Publication: " & .EffectiveResolution & " dpi" 
 Debug.Print "Original Resolution: " & .OriginalResolution & " dpi" 
 
 End With 
 End If 
 Next shpLoop 
Next pgLoop 

```


