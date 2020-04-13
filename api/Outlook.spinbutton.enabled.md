---
title: SpinButton.Enabled Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: d9460bfc-aec4-10b6-fac0-ea9a5977d56c
ms.date: 06/08/2017
localization_priority: Normal
---


# SpinButton.Enabled Property (Outlook Forms Script)

Returns or sets a **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.


## Syntax

_expression_.**Enabled**

_expression_ A variable that represents a **SpinButton** object.


## Remarks

 **True** is the control can receive the focus and respond to user-generated events, and is accessible through code (default). **False** if the user cannot interact with the control by using the mouse, keystrokes, accelerators, or hotkeys. The control is generally still accessible through code.

Use the  **Enabled** property to enable and disable controls. A disabled control appears dimmed, while an enabled control does not. Also, if a control displays a bitmap, the bitmap is dimmed whenever the control is dimmed.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]