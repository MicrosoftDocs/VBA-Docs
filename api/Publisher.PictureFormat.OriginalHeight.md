---
title: PictureFormat.OriginalHeight property (Publisher)
keywords: vbapb10.chm3604774
f1_keywords:
- vbapb10.chm3604774
ms.prod: publisher
api_name:
- Publisher.PictureFormat.OriginalHeight
ms.assetid: 0bf97bb1-d333-a7ed-686c-da2f3cce97c5
ms.date: 06/08/2017
localization_priority: Normal
---


# PictureFormat.OriginalHeight property (Publisher)

Returns a  **Variant** representing the height, in [points](../language/glossary/vbe-glossary.md#point), of the specified linked picture or OLE object. Read-only.


## Syntax

_expression_.**OriginalHeight**

 _expression_ A variable that represents an  **PictureFormat** object.


## Return value

Variant


## Remarks

This property applies only to linked pictures or OLE objects. Returns "Permission Denied" for shapes representing embedded or pasted pictures.

To determine whether a shape represents a linked picture, use either the  **[Type](Publisher.Shape.Type.md)** property of the **[Shape](Publisher.Shape.md)** object, or the **[IsLinked](Publisher.PictureFormat.IsLinked.md)** property of the **[PictureFormat](Publisher.PictureFormat.md)** object.


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
 Debug.Print "Horizontal Scaling: " & .HorizontalScale & "%" 
 Debug.Print "Original Image Height: " & .OriginalHeight & " points" 
 Debug.Print "Height in publication: " & .Height & " points" 
 
 End With 
 End If 
 Next shpLoop 
Next pgLoop 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]