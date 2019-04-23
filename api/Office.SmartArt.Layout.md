---
title: SmartArt.Layout property (Office)
ms.prod: office
api_name:
- Office.SmartArt.Layout
ms.assetid: 5aa76408-9c49-2430-eaea-8893a341b106
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArt.Layout property (Office)

Retrieves or sets the [SmartArt layout](office.smartartlayout.md) associated with the SmartArt graphic. Read/write.


## Syntax

_expression_.**Layout**

_expression_ An expression that returns a **[SmartArt](Office.SmartArt.md)** object.


## Example

The following code sets the SmartArt layout.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.Layout = Application.SmartArtLayouts(1)
```


## See also

- [SmartArt object members](overview/Library-Reference/smartart-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]