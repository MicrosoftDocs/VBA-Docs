---
title: Page.SetTheme method (Visio)
ms.prod: visio
ms.assetid: 5a186f58-9a7a-bd8a-826b-85da75a4d59f
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.SetTheme method (Visio)

Sets the theme for the specified page.


## Syntax

_expression_.**SetTheme** (_varThemeIndex_, _varColorScheme_, _varEffectScheme_, _varConnectorScheme_, _varFontScheme_)

_expression_ A variable that represents a **[Page](Visio.Page.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _varThemeIndex_|Required|**Variant**|The theme to apply.|
| _varColorScheme_|Optional|**Variant**|The color scheme theme component to apply.|
| _varEffectScheme_|Optional|**Variant**|The effect scheme theme component to apply.|
| _varConnectorScheme_|Optional|**Variant**|The connector scheme theme component to apply.|
| _varFontScheme_|Optional|**Variant**|The font scheme theme component to apply.|

## Return value

**VOID**


## Remarks

Possible themes correspond to those displayed in the **Themes** and the **Colors**, **Effects**, and **Connectors** galleries on the **Design** tab of the ribbon. You can specify values for just the first, required parameter, or for any combination of the first parameter and one or more of the other four parameters. 

If you pass a value for the only first parameter, _varThemeIndex_, and you pass nothing for the other four optional parameters, Visio sets all five parameters to the theme value that you specified for the first parameter. 

For example, if you pass "Linear" for the first parameter, Visio sets the color scheme, effect scheme, connector scheme, and font scheme to "Linear" as well. If you pass "Linear" for the first parameter and "Gemstone" for the second parameter, Visio sets the effect scheme, connector scheme, and font scheme to "Linear" but sets the color scheme to "Gemstone" and so on.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]