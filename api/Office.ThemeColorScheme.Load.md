---
title: ThemeColorScheme.Load method (Office)
ms.prod: office
api_name:
- Office.ThemeColorScheme.Load
ms.assetid: 636f14c1-4178-ef12-e22b-4d948719cced
ms.date: 01/25/2019
localization_priority: Normal
---


# ThemeColorScheme.Load method (Office)

Loads the color scheme of a Microsoft Office theme from a file.


## Syntax

_expression_.**Load** (_FileName_)

_expression_ An expression that returns a **[ThemeColorScheme](Office.ThemeColorScheme.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the color theme file.|

## Example

The following example loads a theme color scheme from a file.


```vb
ThemeColorScheme.Load ("C:\myThemeColorScheme.xml") 

```


## See also

- [ThemeColorScheme object members](overview/Library-Reference/themecolorscheme-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]