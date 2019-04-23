---
title: ThemeColor object (Office)
ms.prod: office
api_name:
- Office.ThemeColor
ms.assetid: 357605ea-247d-b151-0286-4e2413658c3f
ms.date: 01/25/2019
localization_priority: Normal
---


# ThemeColor object (Office)

Represents a color in the color scheme of a Microsoft Office theme.


## Example

The following example sets a **ThemeColor** object to the **[msoThemeAccent1](office.msothemecolorschemeindex.md)** constant.


```vb
Dim tcsThemeColorScheme As ThemeColorScheme 
Dim tcThemeColor As ThemeColor 
Set tcThemeColor = tcsThemeColorScheme.Colors(msoThemeAccent1)
```


## See also

- [ThemeColor object members](overview/Library-Reference/themecolor-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]