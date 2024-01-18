---
title: MsoThemeColorSchemeIndex enumeration (Office)
api_name:
- Office.MsoThemeColorSchemeIndex
ms.assetid: a5382465-5552-c131-fad4-d6851f9c0f3e
ms.date: 12/28/2021
ms.localizationpriority: medium
---


# MsoThemeColorSchemeIndex enumeration (Office)

Indicates the color scheme for an Office theme.

|Name|Value|Description|
|:-----|:-----|:-----|
|**msoThemeAccent1**|5|Specifies color scheme Accent 1.|
|**msoThemeAccent2**|6|Specifies color scheme Accent 2.|
|**msoThemeAccent3**|7|Specifies color scheme Accent 3.|
|**msoThemeAccent4**|8|Specifies color scheme Accent 4.|
|**msoThemeAccent5**|9|Specifies color scheme Accent 5.|
|**msoThemeAccent6**|10|Specifies color scheme Accent 6.|
|**msoThemeDark1**|1|Specifies color scheme Dark 1.|
|**msoThemeDark2**|3|Specifies color scheme Dark 2.|
|**msoThemeFollowedHyperlink**|12|Specifies a color scheme for a clicked hyperlink.|
|**msoThemeHyperlink**|11|Specifies a color scheme for a hyperlink.|
|**msoThemeLight1**|2|Specifies color scheme Light 1.|
|**msoThemeLight2**|4|Specifies color scheme Light 2.|

## Remarks

An Office Theme.ThemeColorShceme comprises two light colors, two dark colors, six accent colors and two colors for hyperlinked text. Use this enumeration to set or return the colors for the specified theme. When theme colors are assigned to the ColorFormat object for a shape, they are mapped via the [MsoThemeColorIndex enumeration](/office/vba/api/office.msothemecolorindex). For PowerPoint this mapping takes into account whether the object is present on a light or dark background style for the slide. There are twelve background styles, six light and six dark. For a light background style, an object set to use msoThemeColorBackground1 will be assigned to the Dark 1 color from the theme. For the same object on one of the dark background styles, the Light 1 color is used for the same msoThemeColorBackground1 assignment.

When programmatically assigning a theme color to an object, the MsoThemeColorIndex enumeration should be used, specifically values 13 to 16 for the first four colors of the theme. If values 1 to 4 are used then the Office colour picker UI will not correctly highlight the theme color.

## Example

The following example outputs the Hex color values in BGR format for the twelve colors in the theme for the first slide master in the active presentation, in the order in which they appear in the Office theme editor UI.

```vb
Sub ShowThemeColors()
    With ActivePresentation.Designs(1).SlideMaster.Theme
        Debug.Print Hex(.ThemeColorScheme(msoThemeLight1).RGB)  ' 2
        Debug.Print Hex(.ThemeColorScheme(msoThemeDark1).RGB)   ' 1
        Debug.Print Hex(.ThemeColorScheme(msoThemeLight2).RGB)  ' 4
        Debug.Print Hex(.ThemeColorScheme(msoThemeDark2).RGB)   ' 3
        Debug.Print Hex(.ThemeColorScheme(msoThemeAccent1).RGB)
        Debug.Print Hex(.ThemeColorScheme(msoThemeAccent2).RGB)
        Debug.Print Hex(.ThemeColorScheme(msoThemeAccent3).RGB)
        Debug.Print Hex(.ThemeColorScheme(msoThemeAccent4).RGB)
        Debug.Print Hex(.ThemeColorScheme(msoThemeAccent5).RGB)
        Debug.Print Hex(.ThemeColorScheme(msoThemeHyperlink).RGB)
        Debug.Print Hex(.ThemeColorScheme(msoThemeAccent6).RGB)
        Debug.Print Hex(.ThemeColorScheme(msoThemeFollowedHyperlink).RGB)
    End With
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
