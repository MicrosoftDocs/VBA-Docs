---
title: RGBColor object (PowerPoint)
keywords: vbapp10.chm538000
f1_keywords:
- vbapp10.chm538000
ms.prod: powerpoint
api_name:
- PowerPoint.RGBColor
ms.assetid: 1da5054f-7eaa-37e8-9a5b-d90c790de576
ms.date: 06/08/2017
localization_priority: Normal
---


# RGBColor object (PowerPoint)

Represents a single color in a color scheme.


## Example

Use the [Colors](PowerPoint.ColorScheme.Colors.md)method to return an  **RGBColor** object. You can set an **RGBColor** object to another **RGBColor** object. You can use the [RGB](PowerPoint.RGBColor.RGB.md)property to set or return the explicit red-green-blue value for an  **RGBColor** object, with the exception of the **RGBColor** objects defined by the **ppNotSchemeColor** and **ppSchemeColorMixed** constants. The **RGB** property can be returned, but not set, for these two objects. The following example sets the background color in color scheme one in the active presentation to red and sets the title color to the title color that's defined for color scheme two.


```vb
With ActivePresentation.ColorSchemes

    .Item(1).Colors(ppBackground).RGB = RGB(255, 0, 0)

    .Item(1).Colors(ppTitle) = .Item(2).Colors(ppTitle)

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]