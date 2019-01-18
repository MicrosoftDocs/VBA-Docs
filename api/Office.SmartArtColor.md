---
title: SmartArtColor object (Office)
ms.prod: office
api_name:
- Office.SmartArtColor
ms.assetid: 5aca0209-20d3-c16f-fdfd-184f3464e00b
ms.date: 06/08/2017
localization_priority: Normal
---


# SmartArtColor object (Office)

Chooses the color scheme for the SmartArt diagram.


## Remarks

Simulates the commands on the Microsoft Office Fluent Ribbon user interface on the SmartArt Tools tab, on the Design group, on the Change Colors command.


## Example

The following code sets the color scheme of the Smart Art diagram.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.Color = Application.SmartArtColors(1)
```


## Properties



|Name|
|:-----|
|[Application](Office.SmartArtColor.Application.md)|
|[Category](Office.SmartArtColor.Category.md)|
|[Creator](Office.SmartArtColor.Creator.md)|
|[Description](Office.SmartArtColor.Description.md)|
|[Id](Office.SmartArtColor.Id.md)|
|[Name](Office.SmartArtColor.Name.md)|
|[Parent](Office.SmartArtColor.Parent.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]