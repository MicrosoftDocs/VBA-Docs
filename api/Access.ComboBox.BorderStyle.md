---
title: ComboBox.BorderStyle property (Access)
keywords: vbaac10.chm11408
f1_keywords:
- vbaac10.chm11408
api_name:
- Access.ComboBox.BorderStyle
ms.assetid: ad259ff2-a4b5-14af-3478-3dfc638acab5
ms.date: 02/20/2019
ms.localizationpriority: medium
---


# ComboBox.BorderStyle property (Access)

Specifies how a control's border appears. Read/write **Byte**.


## Syntax

_expression_.**BorderStyle**

_expression_ A variable that represents a **[ComboBox](Access.ComboBox.md)** object.


## Remarks

For controls, the **BorderStyle** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|Transparent|0|(Default only for label, chart, and subreport) Transparent|
|Solid|1|(Default) Solid line|
|Dashes|2|Dashed line|
|Short dashes|3|Dashed line with short dashes|
|Dots|4|Dotted line|
|Sparse dots|5|Dotted line with dots spaced far apart|
|Dash dot|6|Line with a dash-dot combination|
|Dash dot dot|7|Line with a dash-dot-dot combination|

You can set the default for this property by using a control's default control style or the **[DefaultControl](access.form.defaultcontrol.md)** property in Visual Basic.

A control's border style is visible only when its **SpecialEffect** property is set to Flat or Shadowed. If the **SpecialEffect** property is set to something other than Flat or Shadowed, setting the **BorderStyle** property changes the **SpecialEffect** property setting to Flat.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]