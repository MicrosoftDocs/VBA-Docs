---
title: OlkSenderPhoto object (Outlook)
keywords: vbaol11.chm1000498
f1_keywords:
- vbaol11.chm1000498
ms.prod: outlook
api_name:
- Outlook.OlkSenderPhoto
ms.assetid: 07934c3a-404c-7f99-49a8-540701d31cef
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkSenderPhoto object (Outlook)

A control that displays the sender's contact picture for items that can be received via email.


## Remarks

Before you use this control for the first time in the forms designer, add the Microsoft Outlook Sender Photo Control to the control toolbox. You can only add this control to a form region in an Outlook form using the forms designer; you cannot add this control to a Visual Basic  **UserForm** object in the Visual Basic Editor. This control supports Microsoft Windows themes.

If no contact item or contact picture exists for the sender, the control is blank. Right-clicking the control at runtime will display the sender's persona menu, an example of which is shown below.


![Sender menu](../images/olSenderMenu_ZA10120533.gif)



Double-clicking the control will display the contact item inspector.

For more information about Outlook controls, see [Controls in a Custom Form](../outlook/Concepts/Forms/controls-in-a-custom-form.md). For examples of add-ins in C# and Visual Basic .NET that use Outlook controls, see code sample downloads on MSDN. 


## Events



|Name|
|:-----|
|[Change](Outlook.OlkSenderPhoto.Change.md)|
|[Click](Outlook.OlkSenderPhoto.Click.md)|
|[DoubleClick](Outlook.OlkSenderPhoto.DoubleClick.md)|
|[MouseDown](Outlook.OlkSenderPhoto.MouseDown.md)|
|[MouseMove](Outlook.OlkSenderPhoto.MouseMove.md)|
|[MouseUp](Outlook.OlkSenderPhoto.MouseUp.md)|

## Properties



|Name|
|:-----|
|[Enabled](Outlook.OlkSenderPhoto.Enabled.md)|
|[MouseIcon](Outlook.OlkSenderPhoto.MouseIcon.md)|
|[MousePointer](Outlook.OlkSenderPhoto.MousePointer.md)|
|[PreferredHeight](Outlook.OlkSenderPhoto.PreferredHeight.md)|
|[PreferredWidth](Outlook.OlkSenderPhoto.PreferredWidth.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]