---
title: SmartArtQuickStyles Object (Office)
ms.prod: office
api_name:
- Office.SmartArtQuickStyles
ms.assetid: d488ac12-160b-c518-2b56-cc0a3a45c6b7
ms.date: 06/08/2017
---


# SmartArtQuickStyles Object (Office)

Represents a collection of Smart Art quick styles.


## Example

The following code changes the quick style of a Smart Art diagram in Microsoft PowerPoint.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.QuickStyle = Application.SmartArtQuickStyles(i)
```


## Methods



|**Name**|
|:-----|
|[Item](Office.SmartArtQuickStyles.Item.md)|

## Properties



|**Name**|
|:-----|
|[Application](Office.SmartArtQuickStyles.Application.md)|
|[Count](Office.SmartArtQuickStyles.Count.md)|
|[Creator](Office.SmartArtQuickStyles.Creator.md)|
|[Parent](Office.SmartArtQuickStyles.Parent.md)|

## See also





[Object Model Reference](./overview/reference-object-library-reference-for-office.md)
