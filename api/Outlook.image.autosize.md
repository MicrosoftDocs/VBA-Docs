---
title: Image.AutoSize Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 053d8d6f-37d1-98e0-0ef8-e409d9ecaa78
ms.date: 06/08/2017
localization_priority: Normal
---


# Image.AutoSize Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether an object automatically resizes to display its entire contents. Read/write.


## Syntax

_expression_.**AutoSize**

_expression_ A variable that represents an  **Image** object.


## Remarks

 **True** to automatically resize the control to display its entire contents. **False** to keep the size of the control constant; contents are clipped when they exceed the area of the control (default).

For controls without captions, this property specifies whether the control automatically adjusts to display the information stored in the control. In a  **[ComboBox](Outlook.combobox.md)**, for example, setting  **AutoSize** to **True** automatically sets the width of the display area to match the length of the current text.

If you manually change the size of a control while  **AutoSize** is **True**, the manual change overrides the size previously set by  **AutoSize**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]