---
title: BorderArtFormat.StretchPictures property (Publisher)
keywords: vbapb10.chm7602181
f1_keywords:
- vbapb10.chm7602181
ms.prod: publisher
api_name:
- Publisher.BorderArtFormat.StretchPictures
ms.assetid: d3a9c867-111c-a4b1-0e56-6e5ed1e52c8c
ms.date: 06/05/2019
localization_priority: Normal
---


# BorderArtFormat.StretchPictures property (Publisher)

**True** to stretch the picture art making up the specified BorderArt to fit the shape to which it is applied. Read/write **Boolean**. 


## Syntax

_expression_.**StretchPictures**

_expression_ A variable that represents a **[BorderArtFormat](Publisher.BorderArtFormat.md)** object.


## Return value

Boolean


## Remarks

Returns "Permission Denied" if BorderArt has not been applied to the specified object.

Corresponds to the **Don't stretch pictures** and **Stretch pictures to fit** controls in the **BorderArt** dialog box.


## Example

The following example tests for the existence of BorderArt on each shape for each page of the active document. If BorderArt exists, it is set so that it can be stretched.

```vb
Sub StretchBorderArt() 
 Dim anyPage As Page 
 Dim anyShape As Shape 
 
 For Each anyPage in ActiveDocument.Pages 
 For Each anyShape in anyPage.Shapes 
 With anyShape.BorderArt 
 If .Exists = True Then 
 .StretchPictures = True 
 End If 
 End With 
 Next anyShape 
 Next anyPage 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]