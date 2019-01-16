---
title: SmartArtLayout object (Office)
ms.prod: office
api_name:
- Office.SmartArtLayout
ms.assetid: f8d9db83-86f7-4830-096d-5d15368ab6b1
ms.date: 06/08/2017
localization_priority: Normal
---


# SmartArtLayout object (Office)

Represents a Smart Art diagram.


## Remarks

Choices include Basic Block List, Picture Caption List, Vertical Bulleted List, etc.


## Example

The following code changes the diagram style of a Smart Art diagram in Microsoft PowerPoint.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.Layout = Application.SmartArtLayouts(1)
```


## See also


[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)


[SmartArtLayout Object Members](./overview/Library-Reference/smartartlayout-members-office.md)

