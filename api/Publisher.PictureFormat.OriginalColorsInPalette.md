---
title: PictureFormat.OriginalColorsInPalette property (Publisher)
keywords: vbapb10.chm3604771
f1_keywords:
- vbapb10.chm3604771
ms.prod: publisher
api_name:
- Publisher.PictureFormat.OriginalColorsInPalette
ms.assetid: 87c67430-1a5a-47f7-822f-6af8783f73b3
ms.date: 06/13/2019
localization_priority: Normal
---


# PictureFormat.OriginalColorsInPalette property (Publisher)

Returns a **Long** that represents the number of colors in the specified linked picture's palette. Read-only.


## Syntax

_expression_.**OriginalColorsInPalette**

_expression_ A variable that represents a **[PictureFormat](Publisher.PictureFormat.md)** object.


## Return value

Long


## Remarks

This property only applies to linked pictures or OLE objects that are not TrueColor (that is, they contain color data of less than 24 bits per channel). Returns "Permission Denied" for shapes representing embedded or pasted pictures and OLE objects, or linked pictures that are TrueColor.

Use either of the following properties to determine whether a shape represents a linked picture:

- The **[Type](Publisher.Shape.Type.md)** property of the **Shape** object   
- The **[IsLinked](Publisher.PictureFormat.IsLinked.md)** property of the **PictureFormat** object
    
Use the **[OriginalIsTrueColor](Publisher.PictureFormat.OriginalIsTrueColor.md)** property to determine whether a linked picture contains color data of 24 bits per channel or greater.


## Example

The following example returns a list of all pictures in the active publication that are not TrueColor. The number of colors in each picture's palette is returned, and if the picture is linked and the linked picture is not TrueColor, the number of colors in its palette is also returned.

```vb
Sub PictureColorInformation() 
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Or shpLoop.Type = pbPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse Then 
 
 If .IsTrueColor = msoFalse Then 
 Debug.Print .Filename 
 Debug.Print "This picture has " & .ColorsInPalette & " colors." 
 If .IsLinked = msoTrue Then 
 If .OriginalIsTrueColor = msoFalse Then 
 Debug.Print "The linked picture has " & _ 
 .OriginalColorsInPalette & " colors." 
 End If 
 End If 
 End If 
 
 End If 
 End With 
 
 End If 
 Next shpLoop 
Next pgLoop 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]