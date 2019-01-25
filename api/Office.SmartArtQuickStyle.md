---
title: SmartArtQuickStyle object (Office)
ms.prod: office
api_name:
- Office.SmartArtQuickStyle
ms.assetid: e128920b-7adc-71e2-928b-84285f24d574
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArtQuickStyle object (Office)

Represents a SmartArt quick style.


## Example

The following code changes the quick style of a SmartArt diagram in Microsoft PowerPoint.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.QuickStyle = Application.SmartArtQuickStyles(i)
```


## See also

- [SmartArtQuickStyle object members](overview/Library-Reference/smartartquickstyle-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]