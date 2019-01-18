---
title: SmartArt.Layout property (Office)
ms.prod: office
api_name:
- Office.SmartArt.Layout
ms.assetid: 5aa76408-9c49-2430-eaea-8893a341b106
ms.date: 06/08/2017
localization_priority: Normal
---


# SmartArt.Layout property (Office)

Retrieves or sets the Smart Art layout associated with the Smart Art graphic. Read/write


## Syntax

_expression_. `Layout`

 _expression_ An expression that returns a [SmartArt](Office.SmartArt.md) object.


## Example

The following code sets the Smart Art layout.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.Layout = Application.SmartArtLayouts(1)
```


## See also


[SmartArt Object](Office.SmartArt.md)



[SmartArt Object Members](./overview/Library-Reference/smartart-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]