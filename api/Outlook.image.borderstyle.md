---
title: Image.BorderStyle Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: c3b9fb9e-76bb-3ad1-b75a-6acc03b4af9e
ms.date: 06/08/2017
localization_priority: Normal
---


# Image.BorderStyle Property (Outlook Forms Script)

Returns or sets an **Integer** that specifies the type of border of the control. Read/write.


## Syntax

_expression_.**BorderStyle**

_expression_ A variable that represents an **Image** object.


## Remarks

The possible values of  **BorderStyle** are 0 and 1. 0 represents no visible border line, 1 represents a single-line border (default).

 The default value for an **[Image](Outlook.image.md)** is 1 (Single).

You can use either  **BorderStyle** or **[SpecialEffect](Outlook.image.specialeffect.md)** to specify the border for a control, but not both. If you specify a nonzero value for one of these properties, the system sets the value of the other property to zero. For example, if you set **BorderStyle** to 1, the system sets **SpecialEffect** to zero (Flat). If you specify a nonzero value for **SpecialEffect**, the system sets  **BorderStyle** to zero.

 **BorderStyle** uses **[BorderColor](Outlook.image.bordercolor.md)** to define the colors of its borders. To use the **BorderColor** property, the **BorderStyle** property must be set to a value other than 0.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]