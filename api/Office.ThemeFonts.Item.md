---
title: ThemeFonts.Item Method (Office)
ms.prod: office
api_name:
- Office.ThemeFonts.Item
ms.assetid: 09b437dd-9be3-223e-4b81-f83a1d44d53f
ms.date: 06/08/2017
---


# ThemeFonts.Item Method (Office)

Gets one of the three language fonts contained in the  **ThemeFonts** collection.


## Syntax

 _expression_. `Item`( `_Index_` )

 _expression_ An expression that returns a [ThemeFonts](./Office.ThemeFonts.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**MsoFontLanguageIndex**|The index value of the  **ThemeFont** object.|

## Return value

ThemeFont


## Example

The following example sets the font for the body of a document to the Latin theme.


```vb
Dim tTheme As OfficeTheme 
Dim tfThemeFonts As ThemeFonts 
Dim latinMinorFont As ThemeFont 
Set tfThemeFonts = tTheme.ThemeFontScheme.MinorFont 
Set latinMinorFont = tfThemeFonts(msoThemeLatin)
```


## See also


[ThemeFonts Object](Office.ThemeFonts.md)



[ThemeFonts Object Members](./overview/Library-Reference/themefonts-members-office.md)

