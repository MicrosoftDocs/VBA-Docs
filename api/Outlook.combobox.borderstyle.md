---
title: ComboBox.BorderStyle Property (Outlook Forms Script)
keywords: olfm10.chm2000820
f1_keywords:
- olfm10.chm2000820
ms.prod: outlook
ms.assetid: 59caf8ee-9287-362e-1102-c40a9f61bf8d
ms.date: 06/08/2017
localization_priority: Normal
---


# ComboBox.BorderStyle Property (Outlook Forms Script)

Returns or sets an **Integer** that specifies the type of border of the control. Read/write.


## Syntax

_expression_.**BorderStyle**

_expression_ A variable that represents a **ComboBox** object.


## Remarks

The possible values of  **BorderStyle** are 0 and 1. 0 represents no visible border line, 1 represents a single-line border (default).

The default value for a **[ComboBox](Outlook.combobox.md)** is 0 (None).

You can use either  **BorderStyle** or **[SpecialEffect](Outlook.combobox.specialeffect.md)** to specify the border for a control, but not both. If you specify a nonzero value for one of these properties, the system sets the value of the other property to zero. For example, if you set **BorderStyle** to 1, the system sets **SpecialEffect** to zero (Flat). If you specify a nonzero value for **SpecialEffect**, the system sets  **BorderStyle** to zero.

 **BorderStyle** uses **[BorderColor](Outlook.combobox.bordercolor.md)** to define the colors of its borders. To use the **BorderColor** property, the **BorderStyle** property must be set to a value other than 0.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]