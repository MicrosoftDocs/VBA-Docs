---
title: Window.SelectedText property (Visio)
keywords: vis_sdr.chm11651640
f1_keywords:
- vis_sdr.chm11651640
ms.prod: visio
api_name:
- Visio.Window.SelectedText
ms.assetid: 75397f73-192b-7683-2a46-016d9b458879
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.SelectedText property (Visio)

Returns the selected text in the Microsoft Visio drawing window as a  **Characters** object. Read/write.


## Syntax

_expression_. `SelectedText`

_expression_ A variable that represents a **[Window](Visio.Window.md)** object.


## Return value

Characters


## Remarks

The  **SelectedText** property applies only to drawing windows. If you try to access the **SelectedText** property for other types of window, Visio may return an error.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]