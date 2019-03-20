---
title: OptionGroup object (Access)
keywords: vbaac10.chm10894
f1_keywords:
- vbaac10.chm10894
ms.prod: access
api_name:
- Access.OptionGroup
ms.assetid: aa9e5607-7892-9ab2-dabc-822372b23811
ms.date: 03/21/2019
localization_priority: Normal
---


# OptionGroup object (Access)

An option group on a form or report displays a limited set of alternatives. An option group makes selecting a value easy because you can choose the value that you want. Only one option in an option group can be selected at a time.


## Remarks

An option group consists of a group frame and a set of check boxes, toggle buttons, or option buttons.

If an option group is bound to a field, only the group frame itself is bound to the field, not the check boxes, toggle buttons, or option buttons inside the frame. Instead of setting the **ControlSource** property for each control in the option group, you set the **OptionValue** property of each check box, toggle button, or option button to a number that's meaningful for the field to which the group frame is bound. When you select an option in an option group, Microsoft Access sets the value of the field to which the option group is bound to the value of the selected option's **OptionValue** property.

> [!NOTE] 
> The **OptionValue** property is set to a number because the value of an option group can only be a number, not text. Access stores this number in the underlying table. 

An option group can also be set to an expression, or it can be unbound. You can use an unbound option group in a custom dialog box to accept user input and then carry out an action based on that input.


## Events

- [AfterUpdate](Access.OptionGroup.AfterUpdate-event.md)
- [BeforeUpdate](Access.OptionGroup.BeforeUpdate-event.md)
- [Click](Access.OptionGroup.Click.md)
- [DblClick](Access.OptionGroup.DblClick.md)
- [Enter](Access.OptionGroup.Enter.md)
- [Exit](Access.OptionGroup.Exit.md)
- [MouseDown](Access.OptionGroup.MouseDown.md)
- [MouseMove](Access.OptionGroup.MouseMove.md)
- [MouseUp](Access.OptionGroup.MouseUp.md)

## Methods

- [Move](Access.OptionGroup.Move.md)
- [Requery](Access.OptionGroup.Requery.md)
- [SetFocus](Access.OptionGroup.SetFocus.md)
- [SizeToFit](Access.OptionGroup.SizeToFit.md)
- [Undo](Access.OptionGroup.Undo.md)

## Properties

- [AddColon](Access.OptionGroup.AddColon.md)
- [AfterUpdate](Access.OptionGroup.AfterUpdate-property.md)
- [Application](Access.OptionGroup.Application.md)
- [AutoLabel](Access.OptionGroup.AutoLabel.md)
- [BackColor](Access.OptionGroup.BackColor.md)
- [BackShade](Access.OptionGroup.BackShade.md)
- [BackStyle](Access.OptionGroup.BackStyle.md)
- [BackThemeColorIndex](Access.OptionGroup.BackThemeColorIndex.md)
- [BackTint](Access.OptionGroup.BackTint.md)
- [BeforeUpdate](Access.OptionGroup.BeforeUpdate-property.md)
- [BorderColor](Access.OptionGroup.BorderColor.md)
- [BorderShade](Access.OptionGroup.BorderShade.md)
- [BorderStyle](Access.OptionGroup.BorderStyle.md)
- [BorderThemeColorIndex](Access.OptionGroup.BorderThemeColorIndex.md)
- [BorderTint](Access.OptionGroup.BorderTint.md)
- [BorderWidth](Access.OptionGroup.BorderWidth.md)
- [ColumnHidden](Access.OptionGroup.ColumnHidden.md)
- [ColumnOrder](Access.OptionGroup.ColumnOrder.md)
- [ColumnWidth](Access.OptionGroup.ColumnWidth.md)
- [Controls](Access.OptionGroup.Controls.md)
- [ControlSource](Access.OptionGroup.ControlSource.md)
- [ControlTipText](Access.OptionGroup.ControlTipText.md)
- [ControlType](Access.OptionGroup.ControlType.md)
- [DefaultValue](Access.OptionGroup.DefaultValue.md)
- [DisplayWhen](Access.OptionGroup.DisplayWhen.md)
- [Enabled](Access.OptionGroup.Enabled.md)
- [EventProcPrefix](Access.OptionGroup.EventProcPrefix.md)
- [Height](Access.OptionGroup.Height.md)
- [HelpContextId](Access.OptionGroup.HelpContextId.md)
- [HideDuplicates](Access.OptionGroup.HideDuplicates.md)
- [HorizontalAnchor](Access.OptionGroup.HorizontalAnchor.md)
- [InSelection](Access.OptionGroup.InSelection.md)
- [IsVisible](Access.OptionGroup.IsVisible.md)
- [LabelAlign](Access.OptionGroup.LabelAlign.md)
- [LabelX](Access.OptionGroup.LabelX.md)
- [LabelY](Access.OptionGroup.LabelY.md)
- [Left](Access.OptionGroup.Left.md)
- [Locked](Access.OptionGroup.Locked.md)
- [Name](Access.OptionGroup.Name.md)
- [OldBorderStyle](Access.OptionGroup.OldBorderStyle.md)
- [OldValue](Access.OptionGroup.OldValue.md)
- [OnClick](Access.OptionGroup.OnClick.md)
- [OnDblClick](Access.OptionGroup.OnDblClick.md)
- [OnEnter](Access.OptionGroup.OnEnter.md)
- [OnExit](Access.OptionGroup.OnExit.md)
- [OnMouseDown](Access.OptionGroup.OnMouseDown.md)
- [OnMouseMove](Access.OptionGroup.OnMouseMove.md)
- [OnMouseUp](Access.OptionGroup.OnMouseUp.md)
- [Parent](Access.OptionGroup.Parent.md)
- [Properties](Access.OptionGroup.Properties.md)
- [Section](Access.OptionGroup.Section.md)
- [ShortcutMenuBar](Access.OptionGroup.ShortcutMenuBar.md)
- [SpecialEffect](Access.OptionGroup.SpecialEffect.md)
- [StatusBarText](Access.OptionGroup.StatusBarText.md)
- [TabIndex](Access.OptionGroup.TabIndex.md)
- [TabStop](Access.OptionGroup.TabStop.md)
- [Tag](Access.OptionGroup.Tag.md)
- [Top](Access.OptionGroup.Top.md)
- [ValidationRule](Access.OptionGroup.ValidationRule.md)
- [ValidationText](Access.OptionGroup.ValidationText.md)
- [Value](Access.OptionGroup.Value.md)
- [VerticalAnchor](Access.OptionGroup.VerticalAnchor.md)
- [Visible](Access.OptionGroup.Visible.md)
- [Width](Access.OptionGroup.Width.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
