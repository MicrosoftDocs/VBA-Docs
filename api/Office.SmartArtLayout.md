---
title: SmartArtLayout object (Office)
ms.prod: office
api_name:
- Office.SmartArtLayout
ms.assetid: f8d9db83-86f7-4830-096d-5d15368ab6b1
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArtLayout object (Office)

Represents a SmartArt diagram.


## Remarks

Choices include Basic Block list, Picture Caption list, Vertical Bulleted list, etc.


## Example

The following code changes the diagram style of a SmartArt diagram in Microsoft PowerPoint.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.Layout = Application.SmartArtLayouts(1)
```


## See also

- [SmartArtLayout object members](overview/Library-Reference/smartartlayout-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]