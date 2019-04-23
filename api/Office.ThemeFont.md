---
title: ThemeFont object (Office)
ms.prod: office
api_name:
- Office.ThemeFont
ms.assetid: 1a9f1365-c392-3d04-74db-333ac111114a
ms.date: 01/29/2019
localization_priority: Normal
---


# ThemeFont object (Office)

Represents a container for the font schemes of a Microsoft Office theme.


## Example

The following example sets the Headings font scheme in a Microsoft Office theme to a Latin scheme. 


```vb
Dim tTheme As OfficeTheme 
Dim tfThemeFontScheme As ThemeFontScheme 
Dim tfThemeFont As ThemeFont 
Set tfThemeFontScheme = tTheme.ThemeFontScheme 
Set tfThemeFont = tfThemeFontScheme.MajorFont(msoThemeLatin) 

```


## See also

- [ThemeFont object members](overview/Library-Reference/themefont-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]