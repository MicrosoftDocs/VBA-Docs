---
title: ThemeFonts object (Office)
ms.prod: office
api_name:
- Office.ThemeFonts
ms.assetid: 393865af-f008-d26c-5b82-9ae79766e511
ms.date: 01/29/2019
localization_priority: Normal
---


# ThemeFonts object (Office)

Represents a collection of major and minor [fonts](office.themefont.md) in the font scheme of a Microsoft Office theme.


## Example

The following example sets a **ThemeFonts** object to a minor theme font.


```vb
Dim tTheme As OfficeTheme 
Dim tfThemeFonts As ThemeFonts 
Set tfThemeFonts = tTheme.ThemeFontScheme.MinorFont 

```


## See also

- [ThemeFonts object members](overview/Library-Reference/themefonts-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]