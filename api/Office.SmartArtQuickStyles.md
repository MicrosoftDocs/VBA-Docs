---
title: SmartArtQuickStyles object (Office)
ms.prod: office
api_name:
- Office.SmartArtQuickStyles
ms.assetid: d488ac12-160b-c518-2b56-cc0a3a45c6b7
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArtQuickStyles object (Office)

Represents a collection of **[SmartArtQuickStyle](Office.SmartArtQuickStyle.md)** objects.


## Example

The following code changes the quick style of a SmartArt diagram in Microsoft PowerPoint.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.QuickStyle = Application.SmartArtQuickStyles(i)
```


## See also

- [SmartArtQuickStyles object members](overview/Library-Reference/smartartquickstyles-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]