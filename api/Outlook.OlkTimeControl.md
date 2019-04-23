---
title: OlkTimeControl object (Outlook)
keywords: vbaol11.chm1000415
f1_keywords:
- vbaol11.chm1000415
ms.prod: outlook
api_name:
- Outlook.OlkTimeControl
ms.assetid: b23f1741-b920-0caf-d4be-9892d8f2ae07
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkTimeControl object (Outlook)

A control that displays the Outlook time drop-down menu. 


## Remarks

Before you use this control for the first time in the forms designer, add the Microsoft Outlook Time Control to the control toolbox. You can only add this control to a form region in an Outlook form using the forms designer; you cannot add this control to a Visual Basic  **UserForm** object in the Visual Basic Editor.

The time control has several different modes which are exposed via properties on the control. It can be bound to any  **DateTime** property, and can be bound to the same property as a date control to provide the capability to select both date and time.

The following is an example of the time control at runtime. This control supports Microsoft Windows themes.


![Time](../images/olTime_ZA10120552.gif)



If the  **[Click](Outlook.OlkTimeControl.Click.md)** event is implemented but the **[DropButtonClick](Outlook.OlkTimeControl.DropButtonClick.md)** event is not implemented, then clicking the drop button will fire only the **Click** event.

If you bind the time control to the start time or the end time of an appointment item, you must use an add-in to control enabling and disabling of the time control. In particular, when the user sets the appointment as an all-day event, you must use code to disable the time controls for the start time and the end time, and enable the controls only when the user clears this setting.

For more information about Outlook controls, see [Controls in a Custom Form](../outlook/Concepts/Forms/controls-in-a-custom-form.md). For examples of add-ins in C# and Visual Basic .NET that use Outlook controls, see code sample downloads on MSDN. 


## Events



|Name|
|:-----|
|[AfterUpdate](Outlook.OlkTimeControl.AfterUpdate.md)|
|[BeforeUpdate](Outlook.OlkTimeControl.BeforeUpdate.md)|
|[Change](Outlook.OlkTimeControl.Change.md)|
|[Click](Outlook.OlkTimeControl.Click.md)|
|[DoubleClick](Outlook.OlkTimeControl.DoubleClick.md)|
|[DropButtonClick](Outlook.OlkTimeControl.DropButtonClick.md)|
|[Enter](Outlook.OlkTimeControl.Enter.md)|
|[Exit](Outlook.OlkTimeControl.Exit.md)|
|[KeyDown](Outlook.OlkTimeControl.KeyDown.md)|
|[KeyPress](Outlook.OlkTimeControl.KeyPress.md)|
|[KeyUp](Outlook.OlkTimeControl.KeyUp.md)|
|[MouseDown](Outlook.OlkTimeControl.MouseDown.md)|
|[MouseMove](Outlook.OlkTimeControl.MouseMove.md)|
|[MouseUp](Outlook.OlkTimeControl.MouseUp.md)|

## Methods



|Name|
|:-----|
|[DropDown](Outlook.OlkTimeControl.DropDown.md)|

## Properties



|Name|
|:-----|
|[AutoSize](Outlook.OlkTimeControl.AutoSize.md)|
|[AutoWordSelect](Outlook.OlkTimeControl.AutoWordSelect.md)|
|[BackColor](Outlook.OlkTimeControl.BackColor.md)|
|[BackStyle](Outlook.OlkTimeControl.BackStyle.md)|
|[Enabled](Outlook.OlkTimeControl.Enabled.md)|
|[EnterFieldBehavior](Outlook.OlkTimeControl.EnterFieldBehavior.md)|
|[Font](Outlook.OlkTimeControl.Font.md)|
|[ForeColor](Outlook.OlkTimeControl.ForeColor.md)|
|[HideSelection](Outlook.OlkTimeControl.HideSelection.md)|
|[IntervalTime](Outlook.OlkTimeControl.IntervalTime.md)|
|[Locked](Outlook.OlkTimeControl.Locked.md)|
|[MouseIcon](Outlook.OlkTimeControl.MouseIcon.md)|
|[MousePointer](Outlook.OlkTimeControl.MousePointer.md)|
|[ReferenceTime](Outlook.OlkTimeControl.ReferenceTime.md)|
|[Style](Outlook.OlkTimeControl.Style.md)|
|[Text](Outlook.OlkTimeControl.Text.md)|
|[TextAlign](Outlook.OlkTimeControl.TextAlign.md)|
|[Time](Outlook.OlkTimeControl.Time.md)|
|[Value](Outlook.OlkTimeControl.Value.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]