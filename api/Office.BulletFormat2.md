---
title: BulletFormat2 object (Office)
ms.prod: office
api_name:
- Office.BulletFormat2
ms.assetid: ad4c2a05-c34d-fbd4-6b12-3153b94d2c4e
ms.date: 01/02/2019
---


# BulletFormat2 object (Office)

Represents bullet formatting.


## Example

The following example sets the bullet size and color for the paragraphs in shape two on slide one in the active PowerPoint presentation.


```vb
With ActivePresentation.Slides(1).Shapes(2) 
 With .TextFrame.TextRange.ParagraphFormat.BulletFormat2 
 .Visible = True 
 .RelativeSize = 1.25 
 .Character = 169 
 With .Font 
 .Color.RGB = RGB(255, 255, 0) 
 .Name = "Symbol" 
 End With 
 End With 
End With 

```


## Methods

|Name|
|:-----|
|[Picture](Office.BulletFormat2.Picture.md)|

## Properties



|Name|
|:-----|
|[Application](Office.BulletFormat2.Application.md)|
|[Character](Office.BulletFormat2.Character.md)|
|[Creator](Office.BulletFormat2.Creator.md)|
|[Font](Office.BulletFormat2.Font.md)|
|[Number](Office.BulletFormat2.Number.md)|
|[Parent](Office.BulletFormat2.Parent.md)|
|[RelativeSize](Office.BulletFormat2.RelativeSize.md)|
|[StartValue](Office.BulletFormat2.StartValue.md)|
|[Style](Office.BulletFormat2.Style.md)|
|[Type](Office.BulletFormat2.Type.md)|
|[UseTextColor](Office.BulletFormat2.UseTextColor.md)|
|[UseTextFont](Office.BulletFormat2.UseTextFont.md)|
|[Visible](Office.BulletFormat2.Visible.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
