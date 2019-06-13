---
title: PictureFormat.RestoreOriginalColors method (Publisher)
keywords: vbapb10.chm3604800
f1_keywords:
- vbapb10.chm3604800
ms.prod: publisher
api_name:
- Publisher.PictureFormat.RestoreOriginalColors
ms.assetid: 13a0d09f-f809-a1ca-73d9-313ea293d56a
ms.date: 06/13/2019
localization_priority: Normal
---


# PictureFormat.RestoreOriginalColors method (Publisher)

Restores the original colors of a picture that was recolored.


## Syntax

_expression_.**RestoreOriginalColors**

_expression_ A variable that represents a **[PictureFormat](Publisher.PictureFormat.md)** object.


## Remarks

The **RestoreOriginalColors** method corresponds to the **Restore Original Colors** button in the **Recolor Picture** dialog box (**Format** menu > **Picture** > **Recolor**).


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **RestoreOriginalColors** method to restore the original colors of a picture that was recolored by using the **[Recolor](Publisher.PictureFormat.Recolor.md)** method. It recolors the first shape in the **Shapes** collection on the first page of the publication and then restores its original colors.

For this example to work, the recolored shape must be either a picture or an OLE object that represents a picture.

```vb
Public Sub RestoreOriginalColors_Example() 
 
 Dim pubPictureFormat As Publisher.PictureFormat 
 Dim pubShape As Publisher.Shape 
 Dim pubColorFormat As Publisher.ColorFormat 
 
 Set pubShape = ThisDocument.Pages(1).Shapes(1) 
 
 Set pubPictureFormat = pubShape.PictureFormat 
 Set pubColorFormat = pubShape.Fill.BackColor 
 
 pubPictureFormat.Recolor pubColorFormat, msoTrue 
 MsgBox "Picture was recolored." 
 pubPictureFormat.RestoreOriginalColors 
 MsgBox "Original colors in picture were restored." 
 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]