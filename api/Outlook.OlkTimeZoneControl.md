---
title: OlkTimeZoneControl object (Outlook)
keywords: vbaol11.chm1000530
f1_keywords:
- vbaol11.chm1000530
ms.prod: outlook
api_name:
- Outlook.OlkTimeZoneControl
ms.assetid: 2138c4fe-1677-f4f0-1a60-dfac20cc1778
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkTimeZoneControl object (Outlook)

A control that supports a selection from a drop-down list of time zones.


## Remarks

Before you use this control for the first time in the forms designer, add the Microsoft Outlook Time Zone Control to the control toolbox. You can only add this control to a form region in an Outlook form using the Forms Designer; you cannot add this control to a Visual Basic UserForm object in the Visual Basic Editor.

The following is an example of a time zone control. The time zone data can be obtained from the Windows registry key HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones.


![Time zone control](../images/olTimeZoneControl_ZA10174601.gif)



If an appointment item was originally created with a time zone value that no longer exists, the appointment time will be converted to the current local time zone.


## Events



|Name|
|:-----|
|[AfterUpdate](Outlook.OlkTimeZoneControl.AfterUpdate.md)|
|[BeforeUpdate](Outlook.OlkTimeZoneControl.BeforeUpdate.md)|
|[Change](Outlook.OlkTimeZoneControl.Change.md)|
|[Click](Outlook.OlkTimeZoneControl.Click.md)|
|[DoubleClick](Outlook.OlkTimeZoneControl.DoubleClick.md)|
|[DropButtonClick](Outlook.OlkTimeZoneControl.DropButtonClick.md)|
|[Enter](Outlook.OlkTimeZoneControl.Enter.md)|
|[Exit](Outlook.OlkTimeZoneControl.Exit.md)|
|[KeyDown](Outlook.OlkTimeZoneControl.KeyDown.md)|
|[KeyPress](Outlook.OlkTimeZoneControl.KeyPress.md)|
|[KeyUp](Outlook.OlkTimeZoneControl.KeyUp.md)|
|[MouseDown](Outlook.OlkTimeZoneControl.MouseDown.md)|
|[MouseMove](Outlook.OlkTimeZoneControl.MouseMove.md)|
|[MouseUp](Outlook.OlkTimeZoneControl.MouseUp.md)|

## Methods



|Name|
|:-----|
|[DropDown](Outlook.OlkTimeZoneControl.DropDown.md)|

## Properties



|Name|
|:-----|
|[AppointmentTimeField](Outlook.OlkTimeZoneControl.AppointmentTimeField.md)|
|[BorderStyle](Outlook.OlkTimeZoneControl.BorderStyle.md)|
|[Enabled](Outlook.OlkTimeZoneControl.Enabled.md)|
|[Locked](Outlook.OlkTimeZoneControl.Locked.md)|
|[MouseIcon](Outlook.OlkTimeZoneControl.MouseIcon.md)|
|[MousePointer](Outlook.OlkTimeZoneControl.MousePointer.md)|
|[SelectedTimeZoneIndex](Outlook.OlkTimeZoneControl.SelectedTimeZoneIndex.md)|
|[Value](Outlook.OlkTimeZoneControl.Value.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]