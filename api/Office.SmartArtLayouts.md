---
title: SmartArtLayouts object (Office)
ms.prod: office
api_name:
- Office.SmartArtLayouts
ms.assetid: 25e33439-fb5e-01d7-1b85-01884a42ba68
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArtLayouts object (Office)

Represents a collection of SmartArt layout diagrams.


## Remarks

Choices include Basic Block list, Picture Caption list, Vertical Bulleted list, etc.


## Example

The following code changes the diagram style of a SmartArt diagram in Microsoft PowerPoint.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.Layout = Application.SmartArtLayouts(1)
```


## See also

- [SmartArtLayouts object members](overview/Library-Reference/smartartlayouts-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]