---
title: ThemeColorScheme.Colors method (Office)
ms.prod: office
api_name:
- Office.ThemeColorScheme.Colors
ms.assetid: 2ae73cd3-c1b7-1815-5b46-84c349c2535b
ms.date: 01/25/2019
localization_priority: Normal
---


# ThemeColorScheme.Colors method (Office)

Gets an object that represents a color in the color scheme of a Microsoft Office theme.


## Syntax

_expression_.**Colors**(_Index_)

_expression_ An expression that returns a **[ThemeColorScheme](Office.ThemeColorScheme.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**[MsoThemeColorSchemeIndex](office.msothemecolorschemeindex.md)**|The index value of the **ThemeColor** object.|

## Return value

ThemeColor


## Example

In the following example, the **msoThemeAccent1** theme color is set to the color **Red**, and then the scheme is saved to a file.


```vb
Dim tTheme As OfficeTheme 
Dim tcsThemeColorScheme As ThemeColorScheme 
Dim tcThemeColor As ThemeColor 
tcThemeColor.RGB = RGB(255, 0, 0) 
Set tcColorScheme.Colors(msoThemeAccent1) = tcThemeColor 
tcsThemeColorScheme.Save ("C:\myThemeColorScheme.xml") 

```


## See also

- [ThemeColorScheme object members](overview/Library-Reference/themecolorscheme-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
