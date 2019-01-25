---
title: SmartArt.Color property (Office)
ms.prod: office
api_name:
- Office.SmartArt.Color
ms.assetid: 65105010-9780-1b99-ef23-b924300bfccb
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArt.Color property (Office)

Retrieves or sets the [SmartArt color style](office.smartartcolor.md) applied to the SmartArt graphic. Read/write.


## Syntax

_expression_.**Color**

_expression_ An expression that returns a **[SmartArt](Office.SmartArt.md)** object.


## Example

The following code sets the color scheme of the SmartArt diagram.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.Color = Application.SmartArtColors(1)
```


## See also

- [SmartArt object members](overview/Library-Reference/smartart-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]