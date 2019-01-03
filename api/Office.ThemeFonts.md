---
title: ThemeFonts object (Office)
ms.prod: office
api_name:
- Office.ThemeFonts
ms.assetid: 393865af-f008-d26c-5b82-9ae79766e511
ms.date: 06/08/2017
---


# ThemeFonts object (Office)

Represents a collection of major and minor fonts in the font scheme of a Microsoft Office theme.


## Example

The following example sets a  **ThemeFonts** object to a minor theme font.


```vb
Dim tTheme As OfficeTheme 
Dim tfThemeFonts As ThemeFonts 
Set tfThemeFonts = tTheme.ThemeFontScheme.MinorFont 

```


## Methods



|**Name**|
|:-----|
|[Item](Office.ThemeFonts.Item.md)|

## Properties



|**Name**|
|:-----|
|[Application](Office.ThemeFonts.Application.md)|
|[Count](Office.ThemeFonts.Count.md)|
|[Creator](Office.ThemeFonts.Creator.md)|
|[Parent](Office.ThemeFonts.Parent.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
