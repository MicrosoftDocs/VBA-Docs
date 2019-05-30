---
title: Fonts object (PowerPoint)
keywords: vbapp10.chm528000
f1_keywords:
- vbapp10.chm528000
ms.prod: powerpoint
api_name:
- PowerPoint.Fonts
ms.assetid: 1a8f44ea-515f-5eb9-eab5-6204d9b1d5bc
ms.date: 06/08/2017
localization_priority: Normal
---


# Fonts object (PowerPoint)

A collection of all the  **[Font](PowerPoint.Font.md)** objects in the specified presentation.


## Remarks

Each  **Font** object represents a font that's used in the presentation.


## Example

Use the [Fonts](PowerPoint.Presentation.Fonts.md) property to return the **Fonts** collection. The following example displays the number of fonts used in the active presentation.


```vb
MsgBox ActivePresentation.Fonts.Count
```

Use  **Fonts** (_index_), where _index_ is the font's name or index number, to return a single **Font** object. The following example checks to see whether font one in the active presentation is embedded in the presentation.




```vb
If ActivePresentation.Fonts(1).Embedded = True Then 
    MsgBox "Font 1 is embedded"
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]