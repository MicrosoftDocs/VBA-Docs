---
title: NavigationButton.BorderThemeColorIndex property (Access)
keywords: vbaac10.chm14634
f1_keywords:
- vbaac10.chm14634
ms.prod: access
api_name:
- Access.NavigationButton.BorderThemeColorIndex
ms.assetid: 5f03c49f-9d93-d3c6-6f2c-b18cef67ac5d
ms.date: 02/20/2019
localization_priority: Normal
---


# NavigationButton.BorderThemeColorIndex property (Access)

Gets or sets a value that represents a color in the applied color theme associated with the **BorderColor** property of the specified object. Read/write **Long**.


## Syntax

_expression_.**BorderThemeColorIndex**

_expression_ A variable that represents a **[NavigationButton](Access.NavigationButton.md)** object.


## Remarks

The **BorderThemeColorIndex** property contains one of the index values listed in the following table.

|Index value|Description|
|:-----|:-----|
|0|Text 1|
|1|Background 1|
|2|Text 2|
|3|Background 2|
|4|Accent 1|
|5|Accent 2|
|6|Accent 3|
|7|Accent 4|
|8|Accent 5|
|9|Accent 6|
|10|Hyperlink|
|11|Followed Hyperlink|

If no theme is applied, the **BorderThemeColorIndex** property contains -1.

This property is not surfaced in the property sheet.


## Example

The following code example sets the border color to the Text 2 color by setting the **BorderThemeColorIndex** property.


```vb
Me.ctl.BorderThemeColorIndex=2
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]