---
title: OlkDateControl object (Outlook)
keywords: vbaol11.chm1000376
f1_keywords:
- vbaol11.chm1000376
ms.prod: outlook
api_name:
- Outlook.OlkDateControl
ms.assetid: bd0c6bbe-c348-c748-41fe-0cf7ecebcc1e
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkDateControl object (Outlook)

A control that supports the drop-down date picker used in inspectors for task and appointment items to select a date. 


## Remarks

Before you use this control for the first time in the forms designer, add the Microsoft Outlook Date Control to the control toolbox. You can only add this control to a form region in an Outlook form using the forms designer; you cannot add this control to a Visual Basic  **UserForm** object in the Visual Basic Editor.

The following is an example of the date control at runtime. This control supports Microsoft Windows themes.


![Date](../images/olDate_ZA10120280.gif)



This control can bind to any built-in or custom  **DateTime** field. However, the control does not support any date format setting for the field, nor does it support the select range behavior that is available in the appointment inspector.

If the  **[Click](Outlook.OlkDateControl.Click.md)** event is implemented but the **[DropButtonClick](Outlook.OlkDateControl.DropButtonClick.md)** event is not implemented, then clicking the drop button will fire only the **Click** event.

For more information about Outlook controls, see [Controls in a Custom Form](../outlook/Concepts/Forms/controls-in-a-custom-form.md). For examples of add-ins in C# and Visual Basic .NET that use Outlook controls, see code sample downloads on MSDN. 


## Events



|Name|
|:-----|
|[AfterUpdate](Outlook.OlkDateControl.AfterUpdate.md)|
|[BeforeUpdate](Outlook.OlkDateControl.BeforeUpdate.md)|
|[Change](Outlook.OlkDateControl.Change.md)|
|[Click](Outlook.OlkDateControl.Click.md)|
|[DoubleClick](Outlook.OlkDateControl.DoubleClick.md)|
|[DropButtonClick](Outlook.OlkDateControl.DropButtonClick.md)|
|[Enter](Outlook.OlkDateControl.Enter.md)|
|[Exit](Outlook.OlkDateControl.Exit.md)|
|[KeyDown](Outlook.OlkDateControl.KeyDown.md)|
|[KeyPress](Outlook.OlkDateControl.KeyPress.md)|
|[KeyUp](Outlook.OlkDateControl.KeyUp.md)|
|[MouseDown](Outlook.OlkDateControl.MouseDown.md)|
|[MouseMove](Outlook.OlkDateControl.MouseMove.md)|
|[MouseUp](Outlook.OlkDateControl.MouseUp.md)|

## Methods



|Name|
|:-----|
|[DropDown](Outlook.OlkDateControl.DropDown.md)|

## Properties



|Name|
|:-----|
|[AutoSize](Outlook.OlkDateControl.AutoSize.md)|
|[AutoWordSelect](Outlook.OlkDateControl.AutoWordSelect.md)|
|[BackColor](Outlook.OlkDateControl.BackColor.md)|
|[BackStyle](Outlook.OlkDateControl.BackStyle.md)|
|[Date](Outlook.OlkDateControl.Date.md)|
|[Enabled](Outlook.OlkDateControl.Enabled.md)|
|[EnterFieldBehavior](Outlook.OlkDateControl.EnterFieldBehavior.md)|
|[Font](Outlook.OlkDateControl.Font.md)|
|[ForeColor](Outlook.OlkDateControl.ForeColor.md)|
|[HideSelection](Outlook.OlkDateControl.HideSelection.md)|
|[Locked](Outlook.OlkDateControl.Locked.md)|
|[MouseIcon](Outlook.OlkDateControl.MouseIcon.md)|
|[MousePointer](Outlook.OlkDateControl.MousePointer.md)|
|[ShowNoneButton](Outlook.OlkDateControl.ShowNoneButton.md)|
|[Text](Outlook.OlkDateControl.Text.md)|
|[TextAlign](Outlook.OlkDateControl.TextAlign.md)|
|[Value](Outlook.OlkDateControl.Value.md)|

## See also


[OlkDateControl Object Members](overview/Outlook.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
