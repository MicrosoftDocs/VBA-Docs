---
title: SmartArtColors object (Office)
ms.prod: office
api_name:
- Office.SmartArtColors
ms.assetid: a1929517-b1fb-c6fe-b6db-03f7ef1ef894
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArtColors object (Office)

A collection of **[SmartArtColor](Office.SmartArtColor.md)** objects.


## Remarks

Simulates the commands on the [Microsoft Office Fluent Ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md) user interface on the **SmartArt Tools** tab, on the **Design** group, and on the **Change Colors** command.

## Example

The following code sets the color scheme of the SmartArt diagram.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.Color = Application.SmartArtColors(1)
```


## See also

- [SmartArtColors object members](overview/Library-Reference/smartartcolors-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]