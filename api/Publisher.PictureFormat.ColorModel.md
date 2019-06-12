---
title: PictureFormat.ColorModel property (Publisher)
keywords: vbapb10.chm3604753
f1_keywords:
- vbapb10.chm3604753
ms.prod: publisher
api_name:
- Publisher.PictureFormat.ColorModel
ms.assetid: 8e3e259c-943d-c1a9-f090-2ee0f0bb29f2
ms.date: 06/12/2019
localization_priority: Normal
---


# PictureFormat.ColorModel property (Publisher)

Returns a **[PbColorModel](Publisher.PbColorModel.md)** constant that represents the color model of the picture. Read-only.


## Syntax

_expression_.**ColorModel**

_expression_ A variable that represents a **[PictureFormat](Publisher.PictureFormat.md)** object.


## Return value

PbColorModel


## Remarks

The **ColorModel** property value can be one of the **PbColorModel** constants declared in the Microsoft Publisher type library.


## Example

The following example returns a list of the pictures with RGB color mode in the active publication.

```vb
Sub ListRGBPictures() 
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
 For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 
 If shpLoop.Type = pbPicture Or shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse Then 
 If .ColorModel = pbColorModelRGB Then 
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