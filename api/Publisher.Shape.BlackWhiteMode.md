---
title: Shape.BlackWhiteMode Property (Publisher)
keywords: vbapb10.chm2228336
f1_keywords:
- vbapb10.chm2228336
ms.prod: publisher
api_name:
- Publisher.Shape.BlackWhiteMode
ms.assetid: 0a735488-956f-bd3c-ad74-1639780e4e24
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.BlackWhiteMode Property (Publisher)

Returns or sets an  **MsoBlackWhiteMode**constant indicating how the specified shape or shape range appears when the publication is viewed in black-and-white mode. Read/write.


## Syntax

 _expression_. **BlackWhiteMode**

 _expression_ A variable that represents a  **Shape** object.


## Remarks

The  **BlackWhiteMode** property value can be one of the ** [MsoBlackWhiteMode](Office.MsoBlackWhiteMode.md)** constants declared in the Microsoft Office type library.


## Example

This example sets the first shape in the active publication to appear in black-and-white mode. When you view the publication in black-and-white mode, the shape will appear black, regardless of what color it is in color mode.


```vb
ActiveDocument.Pages(1).Shapes(1).BlackWhiteMode = msoBlackWhiteBlack
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]