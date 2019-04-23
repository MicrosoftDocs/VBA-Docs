---
title: OlkComboBox object (Outlook)
keywords: vbaol11.chm1000249
f1_keywords:
- vbaol11.chm1000249
ms.prod: outlook
api_name:
- Outlook.OlkComboBox
ms.assetid: 8d5e2f25-2962-af28-2523-b7b82473ea0a
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkComboBox object (Outlook)

A control that supports the display of a selection from a drop-down list of all choices.


## Remarks

Before you use this control for the first time in the forms designer, add the Microsoft Outlook Combo Box Control to the control toolbox. You can only add this control to a form region in an Outlook form using the forms designer.

The following is an example of a combo box control that has been bound to the Sensitivity field. This control supports Microsoft Windows themes.


![Combo box](../images/olComboBox_ZA10120277.gif)



If the  **[Click](Outlook.OlkComboBox.Click.md)** event is implemented but the **[DropButtonClick](Outlook.OlkComboBox.DropButtonClick.md)** event is not implemented, then clicking the drop button will fire only the **Click** event.

For more information about Outlook controls, see [Controls in a Custom Form](../outlook/Concepts/Forms/controls-in-a-custom-form.md). For examples of add-ins in C# and Visual Basic .NET that use Outlook controls, see code sample downloads on MSDN. 


## Events



|Name|
|:-----|
|[AfterUpdate](Outlook.OlkComboBox.AfterUpdate.md)|
|[BeforeUpdate](Outlook.OlkComboBox.BeforeUpdate.md)|
|[Change](Outlook.OlkComboBox.Change.md)|
|[Click](Outlook.OlkComboBox.Click.md)|
|[DoubleClick](Outlook.OlkComboBox.DoubleClick.md)|
|[DropButtonClick](Outlook.OlkComboBox.DropButtonClick.md)|
|[Enter](Outlook.OlkComboBox.Enter.md)|
|[Exit](Outlook.OlkComboBox.Exit.md)|
|[KeyDown](Outlook.OlkComboBox.KeyDown.md)|
|[KeyPress](Outlook.OlkComboBox.KeyPress.md)|
|[KeyUp](Outlook.OlkComboBox.KeyUp.md)|
|[MouseDown](Outlook.OlkComboBox.MouseDown.md)|
|[MouseMove](Outlook.OlkComboBox.MouseMove.md)|
|[MouseUp](Outlook.OlkComboBox.MouseUp.md)|

## Methods



|Name|
|:-----|
|[AddItem](Outlook.OlkComboBox.AddItem.md)|
|[Clear](Outlook.OlkComboBox.Clear.md)|
|[Copy](Outlook.OlkComboBox.Copy.md)|
|[Cut](Outlook.OlkComboBox.Cut.md)|
|[DropDown](Outlook.OlkComboBox.DropDown.md)|
|[GetItem](Outlook.OlkComboBox.GetItem.md)|
|[Paste](Outlook.OlkComboBox.Paste.md)|
|[RemoveItem](Outlook.OlkComboBox.RemoveItem.md)|
|[SetItem](Outlook.OlkComboBox.SetItem.md)|

## Properties



|Name|
|:-----|
|[AutoSize](Outlook.OlkComboBox.AutoSize.md)|
|[AutoTab](Outlook.OlkComboBox.AutoTab.md)|
|[AutoWordSelect](Outlook.OlkComboBox.AutoWordSelect.md)|
|[BackColor](Outlook.OlkComboBox.BackColor.md)|
|[BorderStyle](Outlook.OlkComboBox.BorderStyle.md)|
|[DragBehavior](Outlook.OlkComboBox.DragBehavior.md)|
|[Enabled](Outlook.OlkComboBox.Enabled.md)|
|[EnterFieldBehavior](Outlook.OlkComboBox.EnterFieldBehavior.md)|
|[Font](Outlook.OlkComboBox.Font.md)|
|[ForeColor](Outlook.OlkComboBox.ForeColor.md)|
|[HideSelection](Outlook.OlkComboBox.HideSelection.md)|
|[ListCount](Outlook.OlkComboBox.ListCount.md)|
|[ListIndex](Outlook.OlkComboBox.ListIndex.md)|
|[Locked](Outlook.OlkComboBox.Locked.md)|
|[MaxLength](Outlook.OlkComboBox.MaxLength.md)|
|[MouseIcon](Outlook.OlkComboBox.MouseIcon.md)|
|[MousePointer](Outlook.OlkComboBox.MousePointer.md)|
|[SelectionMargin](Outlook.OlkComboBox.SelectionMargin.md)|
|[SelLength](Outlook.OlkComboBox.SelLength.md)|
|[SelStart](Outlook.OlkComboBox.SelStart.md)|
|[SelText](Outlook.OlkComboBox.SelText.md)|
|[Style](Outlook.OlkComboBox.Style.md)|
|[Text](Outlook.OlkComboBox.Text.md)|
|[TextAlign](Outlook.OlkComboBox.TextAlign.md)|
|[TopIndex](Outlook.OlkComboBox.TopIndex.md)|
|[Value](Outlook.OlkComboBox.Value.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]