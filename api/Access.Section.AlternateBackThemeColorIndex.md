---
title: Section.AlternateBackThemeColorIndex property (Access)
keywords: vbaac10.chm14607
f1_keywords:
- vbaac10.chm14607
ms.prod: access
api_name:
- Access.Section.AlternateBackThemeColorIndex
ms.assetid: 15ef17dd-06fd-db4a-7253-5d193f2e4b9a
ms.date: 03/23/2019
localization_priority: Normal
---


# Section.AlternateBackThemeColorIndex property (Access)

Gets or sets a value that represents a color in the applied color theme associated with the **AlternateBackColor** property of the section. Read/write **Long**.


## Syntax

_expression_.**AlternateBackThemeColorIndex**

_expression_ A variable that represents a **[Section](Access.Section.md)** object.


## Remarks

The **AlternateBackThemeColorIndex** property uses one of the values listed in the following table.

|Value|Description|
|:-----|:-----|
|0 |Text 1|
|1 |Background 1|
|2|Text 2|
|3 (Default)|Background 2|
|4|Accent 1|
|5|Accent 2|
|6|Accent 3|
|7|Accent 4|
|8|Accent 5|
|9|Accent 6|
|10|Hyperlink|
|11|Followed Hyperlink|

If no theme is applied, the **AlternateBackThemeColorIndex** property contains -1.

This property is not surfaced in the property sheet.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]