---
title: ThemeColorScheme.Load Method (Office)
ms.prod: office
api_name:
- Office.ThemeColorScheme.Load
ms.assetid: 636f14c1-4178-ef12-e22b-4d948719cced
ms.date: 06/08/2017
---


# ThemeColorScheme.Load Method (Office)

Loads the color scheme of a Microsoft Office theme from a file.


## Syntax

 _expression_. `Load`( `_FileName_` )

 _expression_ An expression that returns a [ThemeColorScheme](./Office.ThemeColorScheme.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the color theme file.|

## Example

The following example loads a theme color scheme from a file.


```vb
ThemeColorScheme.Load ("C:\myThemeColorScheme.xml") 

```


## See also


[ThemeColorScheme Object](Office.ThemeColorScheme.md)



[ThemeColorScheme Object Members](./overview/Library-Reference/themecolorscheme-members-office.md)

