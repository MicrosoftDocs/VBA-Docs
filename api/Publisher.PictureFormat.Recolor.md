---
title: PictureFormat.Recolor method (Publisher)
keywords: vbapb10.chm3604793
f1_keywords:
- vbapb10.chm3604793
ms.prod: publisher
api_name:
- Publisher.PictureFormat.Recolor
ms.assetid: 42bc2280-b6d0-862a-7118-38ec1513b9c7
ms.date: 06/13/2019
localization_priority: Normal
---


# PictureFormat.Recolor method (Publisher)

Changes the color of a picture in a publication.


## Syntax

_expression_.**Recolor** (_Color_, _LeaveBlackPartsBlack_)

_expression_ A variable that represents a **[PictureFormat](Publisher.PictureFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Color_ |Required| **ColorFormat**|The color to be used for recoloring.|
|_LeaveBlackPartsBlack_ |Required| **[MsoTriState](office.msotristate.md)** | **True** if all parts of the original picture that were black in color should be left black.|

## Remarks

The **Recolor** method corresponds to the options available in the **Recolor Picture** dialog box (**Format** menu > **Picture** > **Recolor**).


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **Recolor** method to change the color of a picture. It recolors the first shape in the **Shapes** collection on the first page of the publication. After running the code, you can restore the original colors by using the **[RestoreOriginalColors](Publisher.PictureFormat.RestoreOriginalColors.md)** method.

For this example to work, the shape to be recolored must be either a picture or an OLE object that represents a picture.

```vb
Public Sub Recolor_Example() 
 
 Dim pubPictureFormat As Publisher.PictureFormat 
 Dim pubShape As Publisher.Shape 
 Dim pubColorFormat As Publisher.ColorFormat 
 
 Set pubShape = ThisDocument.Pages(1).Shapes(1) 
 
 Set pubPictureFormat = pubShape.PictureFormat 
 Set pubColorFormat = pubShape.Fill.BackColor 
 
 pubPictureFormat.Recolor pubColorFormat, msoTrue 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]