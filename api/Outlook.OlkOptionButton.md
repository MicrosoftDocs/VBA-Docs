---
title: OlkOptionButton object (Outlook)
keywords: vbaol11.chm1000192
f1_keywords:
- vbaol11.chm1000192
ms.prod: outlook
api_name:
- Outlook.OlkOptionButton
ms.assetid: a7aab427-a2f0-a153-f558-c13559610c99
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkOptionButton object (Outlook)

A control that supports a single exclusive selection within a defined group of option button choices.


## Remarks

Before you use this control for the first time in the forms designer, add the Microsoft Outlook Option Button Control to the control toolbox. You can only add this control to a form region in an Outlook form using the forms designer.

The following is an example of an option button control at runtime. This control supports Microsoft Windows themes.


![Option button](../images/olOptionButton_ZA10120824.gif)



Typically more than one option button control is defined in a group. Each option button control in the group provides a single, mutually exclusive choice. Selecting one option button in the group will automatically remove the prior selection of another option button in the same group.

For more information about Outlook controls, see [Controls in a Custom Form](../outlook/Concepts/Forms/controls-in-a-custom-form.md). For examples of add-ins in C# and Visual Basic .NET that use Outlook controls, see code sample downloads on MSDN. 


## Events



|Name|
|:-----|
|[AfterUpdate](Outlook.OlkOptionButton.AfterUpdate.md)|
|[BeforeUpdate](Outlook.OlkOptionButton.BeforeUpdate.md)|
|[Change](Outlook.OlkOptionButton.Change.md)|
|[Click](Outlook.OlkOptionButton.Click.md)|
|[DoubleClick](Outlook.OlkOptionButton.DoubleClick.md)|
|[Enter](Outlook.OlkOptionButton.Enter.md)|
|[Exit](Outlook.OlkOptionButton.Exit.md)|
|[KeyDown](Outlook.OlkOptionButton.KeyDown.md)|
|[KeyPress](Outlook.OlkOptionButton.KeyPress.md)|
|[KeyUp](Outlook.OlkOptionButton.KeyUp.md)|
|[MouseDown](Outlook.OlkOptionButton.MouseDown.md)|
|[MouseMove](Outlook.OlkOptionButton.MouseMove.md)|
|[MouseUp](Outlook.OlkOptionButton.MouseUp.md)|

## Properties



|Name|
|:-----|
|[Accelerator](Outlook.OlkOptionButton.Accelerator.md)|
|[Alignment](Outlook.OlkOptionButton.Alignment.md)|
|[BackColor](Outlook.OlkOptionButton.BackColor.md)|
|[BackStyle](Outlook.OlkOptionButton.BackStyle.md)|
|[Caption](Outlook.OlkOptionButton.Caption.md)|
|[Enabled](Outlook.OlkOptionButton.Enabled.md)|
|[Font](Outlook.OlkOptionButton.Font.md)|
|[ForeColor](Outlook.OlkOptionButton.ForeColor.md)|
|[GroupName](Outlook.OlkOptionButton.GroupName.md)|
|[MouseIcon](Outlook.OlkOptionButton.MouseIcon.md)|
|[MousePointer](Outlook.OlkOptionButton.MousePointer.md)|
|[Value](Outlook.OlkOptionButton.Value.md)|
|[WordWrap](Outlook.OlkOptionButton.WordWrap.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]