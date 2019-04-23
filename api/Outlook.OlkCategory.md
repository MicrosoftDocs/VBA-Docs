---
title: OlkCategory object (Outlook)
keywords: vbaol11.chm1000460
f1_keywords:
- vbaol11.chm1000460
ms.prod: outlook
api_name:
- Outlook.OlkCategory
ms.assetid: f635c0c8-e562-02a2-2a76-25caaee623c0
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkCategory object (Outlook)

A control that displays the selected categories. 


## Remarks

Before you use this control for the first time in the forms designer, add the Microsoft Outlook Category Control to the control toolbox. You can only add this control to a form region in an Outlook form using the forms designer; you cannot add this control to a Visual Basic  **UserForm** object in the Visual Basic Editor.

This control shows both the category name and the category color. Right-clicking this control displays the category selector context menu that allows the user to select categories in standard Outlook forms. If there are no categories to be displayed, this control will resize automatically to a height of zero.

The following is an example of a category control at runtime. This control supports Microsoft Windows themes.


![Category strip](../images/olCategoryStrip_ZA10120276.gif)



For more information about Outlook controls, see [Controls in a Custom Form](../outlook/Concepts/Forms/controls-in-a-custom-form.md). For examples of add-ins in C# and Visual Basic .NET that use Outlook controls, see code sample downloads on MSDN. 


## Events



|Name|
|:-----|
|[Change](Outlook.OlkCategory.Change.md)|
|[Click](Outlook.OlkCategory.Click.md)|
|[DoubleClick](Outlook.OlkCategory.DoubleClick.md)|
|[Enter](Outlook.OlkCategory.Enter.md)|
|[Exit](Outlook.OlkCategory.Exit.md)|
|[KeyDown](Outlook.OlkCategory.KeyDown.md)|
|[KeyPress](Outlook.OlkCategory.KeyPress.md)|
|[KeyUp](Outlook.OlkCategory.KeyUp.md)|
|[MouseDown](Outlook.OlkCategory.MouseDown.md)|
|[MouseMove](Outlook.OlkCategory.MouseMove.md)|
|[MouseUp](Outlook.OlkCategory.MouseUp.md)|

## Properties



|Name|
|:-----|
|[AutoSize](Outlook.OlkCategory.AutoSize.md)|
|[BackColor](Outlook.OlkCategory.BackColor.md)|
|[BackStyle](Outlook.OlkCategory.BackStyle.md)|
|[Enabled](Outlook.OlkCategory.Enabled.md)|
|[ForeColor](Outlook.OlkCategory.ForeColor.md)|
|[MouseIcon](Outlook.OlkCategory.MouseIcon.md)|
|[MousePointer](Outlook.OlkCategory.MousePointer.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]