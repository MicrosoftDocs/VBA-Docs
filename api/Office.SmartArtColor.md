---
title: SmartArtColor object (Office)
ms.prod: office
api_name:
- Office.SmartArtColor
ms.assetid: 5aca0209-20d3-c16f-fdfd-184f3464e00b
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArtColor object (Office)

Chooses the color scheme for the SmartArt diagram.


## Remarks

Simulates the commands on the [Microsoft Office Fluent Ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md) user interface on the **SmartArt Tools** tab, on the **Design** group, and on the **Change Colors** command.


## Example

The following code sets the color scheme of the SmartArt diagram.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.Color = Application.SmartArtColors(1)
```


## See also

- [SmartArtColor object members](overview/Library-Reference/smartartcolor-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]