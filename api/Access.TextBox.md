---
title: TextBox object (Access)
keywords: vbaac10.chm11201
f1_keywords:
- vbaac10.chm11201
ms.prod: access
api_name:
- Access.TextBox
ms.assetid: d74fbe9a-0d40-7d28-956f-a2bfd0cfee45
ms.date: 03/21/2019
localization_priority: Normal
---


# TextBox object (Access)

This object represents a text box control on a form or report. Text boxes are used to display data from a record source, display the results of a calculation, or accept input from a user.

## Remarks

|Control|Tool|
|:-----|:-----|
|![Text box control](../images/t-txtbox_ZA06054010.gif)|![Text box tool](../images/textbox_ZA06044637.gif)|

Text boxes can be either bound or unbound. You use a bound text box to display data from a particular field. You use an unbound text box to display the results of a calculation, or to accept input from a user (as in the following code example).


## Example

The following code example uses a form with a text box to receive user input. The code displays a message when the user inputs data and then presses Enter.

```vb

Private Sub txtValue1_BeforeUpdate(Cancel As Integer)

MsgBox "The Text box is being updated."

End Sub

```

## Events

- [AfterUpdate](Access.TextBox.AfterUpdate-event.md)
- [BeforeUpdate](Access.TextBox.BeforeUpdate-event.md)
- [Change](Access.TextBox.Change.md)
- [Click](Access.TextBox.Click.md)
- [DblClick](Access.TextBox.DblClick.md)
- [Dirty](Access.TextBox.Dirty.md)
- [Enter](Access.TextBox.Enter.md)
- [Exit](Access.TextBox.Exit.md)
- [GotFocus](Access.TextBox.GotFocus.md)
- [KeyDown](Access.TextBox.KeyDown.md)
- [KeyPress](Access.TextBox.KeyPress.md)
- [KeyUp](Access.TextBox.KeyUp.md)
- [LostFocus](Access.TextBox.LostFocus.md)
- [MouseDown](Access.TextBox.MouseDown.md)
- [MouseMove](Access.TextBox.MouseMove.md)
- [MouseUp](Access.TextBox.MouseUp.md)
- [Undo](Access.TextBox.Undo(even).md)

## Methods

- [Move](Access.TextBox.Move.md)
- [Requery](Access.TextBox.Requery.md)
- [SetFocus](Access.TextBox.SetFocus.md)
- [SizeToFit](Access.TextBox.SizeToFit.md)
- [Undo](Access.TextBox.Undo(method).md)

## Properties

- [AddColon](Access.TextBox.AddColon.md)
- [AfterUpdate](Access.TextBox.AfterUpdate-property.md)
- [AllowAutoCorrect](Access.TextBox.AllowAutoCorrect.md)
- [Application](Access.TextBox.Application.md)
- [AsianLineBreak](Access.TextBox.AsianLineBreak.md)
- [AutoLabel](Access.TextBox.AutoLabel.md)
- [AutoTab](Access.TextBox.AutoTab.md)
- [BackColor](Access.TextBox.BackColor.md)
- [BackShade](Access.TextBox.BackShade.md)
- [BackStyle](Access.TextBox.BackStyle.md)
- [BackThemeColorIndex](Access.TextBox.BackThemeColorIndex.md)
- [BackTint](Access.TextBox.BackTint.md)
- [BeforeUpdate](Access.TextBox.BeforeUpdate-property.md)
- [BorderColor](Access.TextBox.BorderColor.md)
- [BorderShade](Access.TextBox.BorderShade.md)
- [BorderStyle](Access.TextBox.BorderStyle.md)
- [BorderThemeColorIndex](Access.TextBox.BorderThemeColorIndex.md)
- [BorderTint](Access.TextBox.BorderTint.md)
- [BorderWidth](Access.TextBox.BorderWidth.md)
- [BottomMargin](Access.TextBox.BottomMargin.md)
- [BottomPadding](Access.TextBox.BottomPadding.md)
- [CanGrow](Access.TextBox.CanGrow.md)
- [CanShrink](Access.TextBox.CanShrink.md)
- [ColumnHidden](Access.TextBox.ColumnHidden.md)
- [ColumnOrder](Access.TextBox.ColumnOrder.md)
- [ColumnWidth](Access.TextBox.ColumnWidth.md)
- [Controls](Access.TextBox.Controls.md)
- [ControlSource](Access.TextBox.ControlSource.md)
- [ControlTipText](Access.TextBox.ControlTipText.md)
- [ControlType](Access.TextBox.ControlType.md)
- [DecimalPlaces](Access.TextBox.DecimalPlaces.md)
- [DefaultValue](Access.TextBox.DefaultValue.md)
- [DisplayAsHyperlink](Access.TextBox.DisplayAsHyperlink.md)
- [DisplayWhen](Access.TextBox.DisplayWhen.md)
- [Enabled](Access.TextBox.Enabled.md)
- [EnterKeyBehavior](Access.TextBox.EnterKeyBehavior.md)
- [EventProcPrefix](Access.TextBox.EventProcPrefix.md)
- [FilterLookup](Access.TextBox.FilterLookup.md)
- [FontBold](Access.TextBox.FontBold.md)
- [FontItalic](Access.TextBox.FontItalic.md)
- [FontName](Access.TextBox.FontName.md)
- [FontSize](Access.TextBox.FontSize.md)
- [FontUnderline](Access.TextBox.FontUnderline.md)
- [FontWeight](Access.TextBox.FontWeight.md)
- [ForeColor](Access.TextBox.ForeColor.md)
- [ForeShade](Access.TextBox.ForeShade.md)
- [ForeThemeColorIndex](Access.TextBox.ForeThemeColorIndex.md)
- [ForeTint](Access.TextBox.ForeTint.md)
- [Format](Access.TextBox.Format.md)
- [FormatConditions](Access.TextBox.FormatConditions.md)
- [FuriganaControl](Access.TextBox.FuriganaControl.md)
- [GridlineColor](Access.TextBox.GridlineColor.md)
- [GridlineShade](Access.TextBox.GridlineShade.md)
- [GridlineStyleBottom](Access.TextBox.GridlineStyleBottom.md)
- [GridlineStyleLeft](Access.TextBox.GridlineStyleLeft.md)
- [GridlineStyleRight](Access.TextBox.GridlineStyleRight.md)
- [GridlineStyleTop](Access.TextBox.GridlineStyleTop.md)
- [GridlineThemeColorIndex](Access.TextBox.GridlineThemeColorIndex.md)
- [GridlineTint](Access.TextBox.GridlineTint.md)
- [GridlineWidthBottom](Access.TextBox.GridlineWidthBottom.md)
- [GridlineWidthLeft](Access.TextBox.GridlineWidthLeft.md)
- [GridlineWidthRight](Access.TextBox.GridlineWidthRight.md)
- [GridlineWidthTop](Access.TextBox.GridlineWidthTop.md)
- [Height](Access.TextBox.Height.md)
- [HelpContextId](Access.TextBox.HelpContextId.md)
- [HideDuplicates](Access.TextBox.HideDuplicates.md)
- [HorizontalAnchor](Access.TextBox.HorizontalAnchor.md)
- [Hyperlink](Access.TextBox.Hyperlink.md)
- [IMEHold](Access.TextBox.IMEHold.md)
- [IMEMode](Access.TextBox.IMEMode.md)
- [IMESentenceMode](Access.TextBox.IMESentenceMode.md)
- [InputMask](Access.TextBox.InputMask.md)
- [InSelection](Access.TextBox.InSelection.md)
- [IsHyperlink](Access.TextBox.IsHyperlink.md)
- [IsVisible](Access.TextBox.IsVisible.md)
- [KeyboardLanguage](Access.TextBox.KeyboardLanguage.md)
- [LabelAlign](Access.TextBox.LabelAlign.md)
- [LabelX](Access.TextBox.LabelX.md)
- [LabelY](Access.TextBox.LabelY.md)
- [Layout](Access.TextBox.Layout.md)
- [LayoutID](Access.TextBox.LayoutID.md)
- [Left](Access.TextBox.Left.md)
- [LeftMargin](Access.TextBox.LeftMargin.md)
- [LeftPadding](Access.TextBox.LeftPadding.md)
- [LineSpacing](Access.TextBox.LineSpacing.md)
- [Locked](Access.TextBox.Locked.md)
- [Name](Access.TextBox.Name.md)
- [NumeralShapes](Access.TextBox.NumeralShapes.md)
- [OldBorderStyle](Access.TextBox.OldBorderStyle.md)
- [OldValue](Access.TextBox.OldValue.md)
- [OnChange](Access.TextBox.OnChange.md)
- [OnClick](Access.TextBox.OnClick.md)
- [OnDblClick](Access.TextBox.OnDblClick.md)
- [OnDirty](Access.TextBox.OnDirty.md)
- [OnEnter](Access.TextBox.OnEnter.md)
- [OnExit](Access.TextBox.OnExit.md)
- [OnGotFocus](Access.TextBox.OnGotFocus.md)
- [OnKeyDown](Access.TextBox.OnKeyDown.md)
- [OnKeyPress](Access.TextBox.OnKeyPress.md)
- [OnKeyUp](Access.TextBox.OnKeyUp.md)
- [OnLostFocus](Access.TextBox.OnLostFocus.md)
- [OnMouseDown](Access.TextBox.OnMouseDown.md)
- [OnMouseMove](Access.TextBox.OnMouseMove.md)
- [OnMouseUp](Access.TextBox.OnMouseUp.md)
- [OnUndo](Access.TextBox.OnUndo.md)
- [Parent](Access.TextBox.Parent.md)
- [PostalAddress](Access.TextBox.PostalAddress.md)
- [Properties](Access.TextBox.Properties.md)
- [ReadingOrder](Access.TextBox.ReadingOrder.md)
- [RightMargin](Access.TextBox.RightMargin.md)
- [RightPadding](Access.TextBox.RightPadding.md)
- [RunningSum](Access.TextBox.RunningSum.md)
- [ScrollBarAlign](Access.TextBox.ScrollBarAlign.md)
- [ScrollBars](Access.TextBox.ScrollBars.md)
- [Section](Access.TextBox.Section.md)
- [SelLength](Access.TextBox.SelLength.md)
- [SelStart](Access.TextBox.SelStart.md)
- [SelText](Access.TextBox.SelText.md)
- [ShortcutMenuBar](Access.TextBox.ShortcutMenuBar.md)
- [ShowDatePicker](Access.TextBox.ShowDatePicker.md)
- [SmartTags](Access.TextBox.SmartTags.md)
- [SpecialEffect](Access.TextBox.SpecialEffect.md)
- [StatusBarText](Access.TextBox.StatusBarText.md)
- [TabIndex](Access.TextBox.TabIndex.md)
- [TabStop](Access.TextBox.TabStop.md)
- [Tag](Access.TextBox.Tag.md)
- [Text](Access.TextBox.Text.md)
- [TextAlign](Access.TextBox.TextAlign.md)
- [TextFormat](Access.TextBox.TextFormat.md)
- [ThemeFontIndex](Access.TextBox.ThemeFontIndex.md)
- [Top](Access.TextBox.Top.md)
- [TopMargin](Access.TextBox.TopMargin.md)
- [TopPadding](Access.TextBox.TopPadding.md)
- [ValidationRule](Access.TextBox.ValidationRule.md)
- [ValidationText](Access.TextBox.ValidationText.md)
- [Value](Access.TextBox.Value.md)
- [Vertical](Access.TextBox.Vertical.md)
- [VerticalAnchor](Access.TextBox.VerticalAnchor.md)
- [Visible](Access.TextBox.Visible.md)
- [Width](Access.TextBox.Width.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
