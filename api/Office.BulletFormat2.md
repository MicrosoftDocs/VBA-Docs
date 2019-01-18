---
title: BulletFormat2 object (Office)
ms.prod: office
api_name:
- Office.BulletFormat2
ms.assetid: ad4c2a05-c34d-fbd4-6b12-3153b94d2c4e
ms.date: 01/02/2019
localization_priority: Normal
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


## See also

- [BulletFormat2 members](overview/library-reference/bulletformat2-members-office.md)
- [Object Model Reference](overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]