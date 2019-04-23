---
title: GlowFormat object (Office)
ms.prod: office
api_name:
- Office.GlowFormat
ms.assetid: b89e2245-e3a4-4a8c-cd4f-86396ad71a5b
ms.date: 01/16/2019
localization_priority: Normal
---


# GlowFormat object (Office)

Represents a glow effect around an Office graphic.


## Example

This example applies glow to the text in the second shape on the second slide in a PowerPoint presentation:


```vb
With ActivePresentation.Slides(2).Shapes(2) 
 .Text.Font.Glowformat = msoGlowType2 
End With 

```


## See also

- [GlowFormat object members](overview/library-reference/glowformat-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]