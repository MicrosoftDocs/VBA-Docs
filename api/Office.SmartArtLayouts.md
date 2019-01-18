---
title: SmartArtLayouts object (Office)
ms.prod: office
api_name:
- Office.SmartArtLayouts
ms.assetid: 25e33439-fb5e-01d7-1b85-01884a42ba68
ms.date: 06/08/2017
localization_priority: Normal
---


# SmartArtLayouts object (Office)

Represents a collection of Smart Art layout diagrams.


## Remarks

Choices include Basic Block List, Picture Caption List, Vertical Bulleted List, etc.


## Example

The following code changes the diagram style of a Smart Art diagram in Microsoft PowerPoint.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.Layout = Application.SmartArtLayouts(1)
```


## See also


[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)

[SmartArtLayouts Object Members](./overview/Library-Reference/smartartlayouts-members-office.md)

