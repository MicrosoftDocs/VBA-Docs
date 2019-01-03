---
title: SmartArt.Color property (Office)
ms.prod: office
api_name:
- Office.SmartArt.Color
ms.assetid: 65105010-9780-1b99-ef23-b924300bfccb
ms.date: 06/08/2017
---


# SmartArt.Color property (Office)

Retrieves or sets the Smart Art color style applied to the Smart Art graphic. Read/write


## Syntax

_expression_. `Color`

 _expression_ An expression that returns a [SmartArt](Office.SmartArt.md) object.


## Example

The following code sets the color scheme of the Smart Art diagram.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.Color = Application.SmartArtColors(1)
```


## See also


[SmartArt Object](Office.SmartArt.md)



[SmartArt Object Members](./overview/Library-Reference/smartart-members-office.md)

