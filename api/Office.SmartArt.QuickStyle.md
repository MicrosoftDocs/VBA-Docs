---
title: SmartArt.QuickStyle property (Office)
ms.prod: office
api_name:
- Office.SmartArt.QuickStyle
ms.assetid: 7f3f8f2f-0b41-4638-2ecc-dd6650f4e98e
ms.date: 06/08/2017
localization_priority: Normal
---


# SmartArt.QuickStyle property (Office)

Retrieves or sets the SmartArt quick style applied to the SmartArt graphic. Read/write


## Syntax

_expression_. `QuickStyle`

 _expression_ An expression that returns a [SmartArt](Office.SmartArt.md) object.


## Example

The following code changes the quick style of Smart Art in Microsoft PowerPoint.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.QuickStyle = Application.SmartArtQuickStyles(i)
```


## See also


[SmartArt Object](Office.SmartArt.md)



[SmartArt Object Members](./overview/Library-Reference/smartart-members-office.md)

