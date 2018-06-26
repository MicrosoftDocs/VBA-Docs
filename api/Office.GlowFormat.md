---
title: GlowFormat Object (Office)
ms.prod: office
api_name:
- Office.GlowFormat
ms.assetid: b89e2245-e3a4-4a8c-cd4f-86396ad71a5b
ms.date: 06/08/2017
---


# GlowFormat Object (Office)

Represents a glow effect around an Office graphic.


## Example

This example applies glow to the text in the second shape on the second slide in a PowerPoint presentation:


```vb
With ActivePresentation.Slides(2).Shapes(2) 
 .Text.Font.Glowformat = msoGlowType2 
End With 

```


## Properties



|**Name**|
|:-----|
|[Application](Office.GlowFormat.Application.md)|
|[Color](Office.GlowFormat.Color.md)|
|[Creator](Office.GlowFormat.Creator.md)|
|[Radius](Office.GlowFormat.Radius.md)|
|[Transparency](Office.GlowFormat.Transparency.md)|

## See also


#### Other resources


[Object Model Reference](./overview/reference-object-library-reference-for-office.md)
