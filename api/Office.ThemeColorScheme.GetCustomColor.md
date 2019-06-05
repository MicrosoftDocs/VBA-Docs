---
title: ThemeColorScheme.GetCustomColor method (Office)
ms.prod: office
api_name:
- Office.ThemeColorScheme.GetCustomColor
ms.assetid: 67ac156e-19ab-245e-b6f8-03514f802acb
ms.date: 01/25/2019
localization_priority: Normal
---


# ThemeColorScheme.GetCustomColor method (Office)

Gets a value that represents a color in the color scheme of a Microsoft Office theme. 


## Syntax

_expression_.**GetCustomColor**(_Name_)

_expression_ An expression that returns a **[ThemeColorScheme](Office.ThemeColorScheme.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the custom color.|

## Return value

**[MsoThemeColorSchemeIndex](office.msothemecolorschemeindex.md)**


## Remarks

If the named custom color doesn't exist, an error is generated.


## Example

The following example creates a variable representing the color scheme in an Office theme, and then creates another variable containing a custom color. This custom color can then be combined with other colors to define the theme.


```vb
Dim tTheme As OfficeTheme 
Dim tcsThemeColorScheme As ThemeColorScheme 
Dim csCustomColor As MsoThemeColorSchemeIndex 
Set tcsThemeColorScheme = tTheme.ThemeColorScheme 
csCustomColor = tcsThemeColorScheme.GetCustomColor("CheerfulColor") 

```


## See also

- [ThemeColorScheme object members](overview/Library-Reference/themecolorscheme-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]