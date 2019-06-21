---
title: UIObject.DisplayTooltips property (Visio)
keywords: vis_sdr.chm14913415
f1_keywords:
- vis_sdr.chm14913415
ms.prod: visio
api_name:
- Visio.UIObject.DisplayTooltips
ms.assetid: 601cf4a4-5afe-1835-4afb-d21f801b93ce
ms.date: 06/08/2017
localization_priority: Normal
---


# UIObject.DisplayTooltips property (Visio)

Determines whether feature descriptions are shown in ScreenTips. Read/write.


## Syntax

_expression_. `DisplayTooltips`

_expression_ A variable that represents a **[UIObject](Visio.UIObject.md)** object.


## Return value

Boolean


## Remarks

It does not matter which  **UIObject** object you use when getting or setting this property. The property affects the entire application.

This property setting corresponds to the  **ScreenTip style** setting on the **General** tab in the **Visio Options** dialog box (click the **File** tab, and then click **Options**), and is shared between Visio and all Microsoft Office applications.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]