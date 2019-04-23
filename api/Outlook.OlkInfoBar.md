---
title: OlkInfoBar object (Outlook)
keywords: vbaol11.chm1000304
f1_keywords:
- vbaol11.chm1000304
ms.prod: outlook
api_name:
- Outlook.OlkInfoBar
ms.assetid: 1aec19db-d28b-ef9b-3227-45aa4a296de6
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkInfoBar object (Outlook)

A control that provides an area to display specific information on a custom form.


## Remarks

Before you use this control for the first time in the forms designer, add the Microsoft Outlook InfoBar Control to the control toolbox. You can only add this control to a form region in an Outlook form using the forms designer; you cannot add this control to a Visual Basic  **UserForm** object in the Visual Basic Editor.

The following is an example of this control at runtime. This control supports Microsoft Windows themes.


![Information bar](../images/olInfoBar_ZA10119648.gif)



If there is no information to display, this control will automatically resize to a height of zero.

You can specify only the placement of the control, as there are no configurable options or settings other than its position.

For more information about Outlook controls, see [Controls in a Custom Form](../outlook/Concepts/Forms/controls-in-a-custom-form.md). For examples of add-ins in C# and Visual Basic .NET that use Outlook controls, see code sample downloads on MSDN. 


## Events



|Name|
|:-----|
|[Click](Outlook.OlkInfoBar.Click.md)|
|[DoubleClick](Outlook.OlkInfoBar.DoubleClick.md)|
|[MouseDown](Outlook.OlkInfoBar.MouseDown.md)|
|[MouseMove](Outlook.OlkInfoBar.MouseMove.md)|
|[MouseUp](Outlook.OlkInfoBar.MouseUp.md)|

## Properties



|Name|
|:-----|
|[MouseIcon](Outlook.OlkInfoBar.MouseIcon.md)|
|[MousePointer](Outlook.OlkInfoBar.MousePointer.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]