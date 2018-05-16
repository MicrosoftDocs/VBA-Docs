---
title: WebBrowserControl.BorderStyle Property (Access)
keywords: vbaac10.chm14375
f1_keywords:
- vbaac10.chm14375
ms.prod: access
api_name:
- Access.WebBrowserControl.BorderStyle
ms.assetid: 3f53cca7-0cfd-2e70-5a20-ce4dbdf166da
ms.date: 06/08/2017
---


# WebBrowserControl.BorderStyle Property (Access)

Specifies how a control's border appears.Read/write  **Byte**.


## Syntax

 _expression_. **BorderStyle**

 _expression_ A variable that represents a **WebBrowserControl** object.


## Remarks

For controls, the  **BorderStyle** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Transparent|0|(Default only for label, chart, and subreport) Transparent|
|Solid|1|(Default) Solid line|
|Dashes|2|Dashed line|
|Short dashes|3|Dashed line with short dashes|
|Dots|4|Dotted line|
|Sparse dots|5|Dotted line with dots spaced far apart|
|Dash dot|6|Line with a dash-dot combination|
|Dash dot dot|7|Line with a dash-dot-dot combination|
|Double solid|8|Double solid lines|
You can set the default for this property by using a control's default control style or the  **DefaultControl** property in Visual Basic.

A control's border style is visible only when its  **SpecialEffect** property is set to Flat or Shadowed. If the **SpecialEffect** property is set to something other than Flat or Shadowed, setting the **BorderStyle** property changes the **SpecialEffect** property setting to Flat.


## See also


#### Concepts


[WebBrowserControl Object](Access.WebBrowserControl.md)

