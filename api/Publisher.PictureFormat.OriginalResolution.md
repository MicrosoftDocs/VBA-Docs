---
title: PictureFormat.OriginalResolution property (Publisher)
keywords: vbapb10.chm3604776
f1_keywords:
- vbapb10.chm3604776
ms.prod: publisher
api_name:
- Publisher.PictureFormat.OriginalResolution
ms.assetid: 0cb7ee4e-3eb8-baee-6535-d936e3c5f05c
ms.date: 06/13/2019
localization_priority: Normal
---


# PictureFormat.OriginalResolution property (Publisher)

Returns a **Long** that represents, in dots per inch (dpi), the resolution at which the linked picture was originally scanned. Read-only.


## Syntax

_expression_.**OriginalResolution**

_expression_ A variable that represents a **[PictureFormat](Publisher.PictureFormat.md)** object.


## Return value

Long


## Remarks

This property only applies to linked pictures. Returns "Permission Denied" for shapes representing embedded or pasted pictures.

Use either of the following properties to determine whether a shape represents a linked picture:

- The **[Type](Publisher.Shape.Type.md)** property of the **Shape** object   
- The **[IsLinked](Publisher.PictureFormat.IsLinked.md)** property of the **PictureFormat** object

Use the **[EffectiveResolution](Publisher.PictureFormat.EffectiveResolution.md)** property to determine the resolution at which the picture or OLE object prints in the specified document.


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

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]