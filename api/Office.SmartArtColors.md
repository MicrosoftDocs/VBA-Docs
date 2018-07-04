---
title: SmartArtColors Object (Office)
ms.prod: office
api_name:
- Office.SmartArtColors
ms.assetid: a1929517-b1fb-c6fe-b6db-03f7ef1ef894
ms.date: 06/08/2017
---


# SmartArtColors Object (Office)

A collection of SmartArt color styles.


## Remarks

Simulates the commands on the Microsoft Office Fluent Ribbon user interface on the SmartArt Tools, on the Design group, on the Change Colors command.


## Example

The following code sets the color scheme of the Smart Art diagram.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.Color = Application.SmartArtColors(1)
```


## Methods



|**Name**|
|:-----|
|[Item](Office.SmartArtColors.Item.md)|

## Properties



|**Name**|
|:-----|
|[Application](Office.SmartArtColors.Application.md)|
|[Count](Office.SmartArtColors.Count.md)|
|[Creator](Office.SmartArtColors.Creator.md)|
|[Parent](Office.SmartArtColors.Parent.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
