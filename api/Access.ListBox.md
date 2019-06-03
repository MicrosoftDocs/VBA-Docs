---
title: ListBox object (Access)
keywords: vbaac10.chm11354
f1_keywords:
- vbaac10.chm11354
ms.prod: access
api_name:
- Access.ListBox
ms.assetid: 6bc00755-34e7-4fc2-8e72-40dae2010dd8
ms.date: 03/21/2019
localization_priority: Normal
---


# ListBox object (Access)

This object corresponds to a list box control. The list box control displays a list of values or alternatives.


## Remarks

|Control|Tool|
|:-----|:-----|
|![List box control](../images/t-lstbox_ZA06053984.gif)|![List box tool](../images/listbox_ZA06044481.gif)|

In many cases, it's quicker and easier to select a value from a list than to remember a value to type. A list of choices also helps ensure that the value that's entered in a field is correct.

The list in a list box consists of rows of data. Rows can have one or more columns, which can appear with or without headings, as shown in the following diagram.

![Multi-column list box](../images/cfrmlst2_ZA06047456.gif)

If a multiple-column list box is bound, Microsoft Access stores the values from one of the columns.

You can use an unbound list box to store a value that you can use with another control. For example, you could use an unbound list box to limit the values in another list box or in a custom dialog box. You could also use an unbound list box to find a record based on the value that you select in the list box.

If you don't have room on your form to display a list box, or if you want to be able to type new values as well as select values from a list, use a combo box instead of a list box.

## Example

This example demonstrates how to filter the contents of a list box while you are typing in a text box.

In this example, a list box named **ColorID** displays a list of colors stored in the **Colors** table. As you type in the **FilterBy** text box, the items in **ColorID** are filtered dynamically.

To do this, use the **Change** event of the text box to build a SQL statement that will serve as the new RowSource of the list box.

```vb
Private Sub FilterBy_Change()

    Dim sql As String
    
    'This will match any entry in the list that begins with what the user 
    'has typed in the FilterBy control
    sql = "SELECT ColorID, ColorName FROM Colors WHERE ColorName Like '" & Me.FilterBy.Text & "*' ORDER BY ColorName"
    
    'If you want to match any part of the string then add wildcard (*) before
    'the FilterBy.Text, too:
    'sql = "SELECT ColorID, ColorName FROM Colors WHERE ColorName Like '*" & Me.FilterBy.Text & "*' ORDER BY ColorName"
    
    Me.ColorID.RowSource = sql
    
End Sub
```


## Events

- [AfterUpdate](Access.ListBox.AfterUpdate-event.md)
- [BeforeUpdate](Access.ListBox.BeforeUpdate-event.md)
- [Click](Access.ListBox.Click.md)
- [DblClick](Access.ListBox.DblClick.md)
- [Enter](Access.ListBox.Enter.md)
- [Exit](Access.ListBox.Exit.md)
- [GotFocus](Access.ListBox.GotFocus.md)
- [KeyDown](Access.ListBox.KeyDown.md)
- [KeyPress](Access.ListBox.KeyPress.md)
- [KeyUp](Access.ListBox.KeyUp.md)
- [LostFocus](Access.ListBox.LostFocus.md)
- [MouseDown](Access.ListBox.MouseDown.md)
- [MouseMove](Access.ListBox.MouseMove.md)
- [MouseUp](Access.ListBox.MouseUp.md)

## Methods

- [AddItem](Access.ListBox.AddItem.md)
- [Move](Access.ListBox.Move.md)
- [RemoveItem](Access.ListBox.RemoveItem.md)
- [Requery](Access.ListBox.Requery.md)
- [SetFocus](Access.ListBox.SetFocus.md)
- [SizeToFit](Access.ListBox.SizeToFit.md)
- [Undo](Access.ListBox.Undo.md)

## Properties

- [AddColon](Access.ListBox.AddColon.md)
- [AfterUpdate](Access.ListBox.AfterUpdate-property.md)
- [AllowValueListEdits](Access.ListBox.AllowValueListEdits.md)
- [Application](Access.ListBox.Application.md)
- [AutoLabel](Access.ListBox.AutoLabel.md)
- [BackColor](Access.ListBox.BackColor.md)
- [BackShade](Access.ListBox.BackShade.md)
- [BackThemeColorIndex](Access.ListBox.BackThemeColorIndex.md)
- [BackTint](Access.ListBox.BackTint.md)
- [BeforeUpdate](Access.ListBox.BeforeUpdate-property.md)
- [BorderColor](Access.ListBox.BorderColor.md)
- [BorderShade](Access.ListBox.BorderShade.md)
- [BorderStyle](Access.ListBox.BorderStyle.md)
- [BorderThemeColorIndex](Access.ListBox.BorderThemeColorIndex.md)
- [BorderTint](Access.ListBox.BorderTint.md)
- [BorderWidth](Access.ListBox.BorderWidth.md)
- [BottomPadding](Access.ListBox.BottomPadding.md)
- [BoundColumn](Access.ListBox.BoundColumn.md)
- [Column](Access.ListBox.Column.md)
- [ColumnCount](Access.ListBox.ColumnCount.md)
- [ColumnHeads](Access.ListBox.ColumnHeads.md)
- [ColumnHidden](Access.ListBox.ColumnHidden.md)
- [ColumnOrder](Access.ListBox.ColumnOrder.md)
- [ColumnWidth](Access.ListBox.ColumnWidth.md)
- [ColumnWidths](Access.ListBox.ColumnWidths.md)
- [Controls](Access.ListBox.Controls.md)
- [ControlSource](Access.ListBox.ControlSource.md)
- [ControlTipText](Access.ListBox.ControlTipText.md)
- [ControlType](Access.ListBox.ControlType.md)
- [DefaultValue](Access.ListBox.DefaultValue.md)
- [DisplayWhen](Access.ListBox.DisplayWhen.md)
- [Enabled](Access.ListBox.Enabled.md)
- [EventProcPrefix](Access.ListBox.EventProcPrefix.md)
- [FontBold](Access.ListBox.FontBold.md)
- [FontItalic](Access.ListBox.FontItalic.md)
- [FontName](Access.ListBox.FontName.md)
- [FontSize](Access.ListBox.FontSize.md)
- [FontUnderline](Access.ListBox.FontUnderline.md)
- [FontWeight](Access.ListBox.FontWeight.md)
- [ForeColor](Access.ListBox.ForeColor.md)
- [ForeShade](Access.ListBox.ForeShade.md)
- [ForeThemeColorIndex](Access.ListBox.ForeThemeColorIndex.md)
- [ForeTint](Access.ListBox.ForeTint.md)
- [GridlineColor](Access.ListBox.GridlineColor.md)
- [GridlineShade](Access.ListBox.GridlineShade.md)
- [GridlineStyleBottom](Access.ListBox.GridlineStyleBottom.md)
- [GridlineStyleLeft](Access.ListBox.GridlineStyleLeft.md)
- [GridlineStyleRight](Access.ListBox.GridlineStyleRight.md)
- [GridlineStyleTop](Access.ListBox.GridlineStyleTop.md)
- [GridlineThemeColorIndex](Access.ListBox.GridlineThemeColorIndex.md)
- [GridlineTint](Access.ListBox.GridlineTint.md)
- [GridlineWidthBottom](Access.ListBox.GridlineWidthBottom.md)
- [GridlineWidthLeft](Access.ListBox.GridlineWidthLeft.md)
- [GridlineWidthRight](Access.ListBox.GridlineWidthRight.md)
- [GridlineWidthTop](Access.ListBox.GridlineWidthTop.md)
- [Height](Access.ListBox.Height.md)
- [HelpContextId](Access.ListBox.HelpContextId.md)
- [HideDuplicates](Access.ListBox.HideDuplicates.md)
- [HorizontalAnchor](Access.ListBox.HorizontalAnchor.md)
- [Hyperlink](Access.ListBox.Hyperlink.md)
- [IMEHold](Access.ListBox.IMEHold.md)
- [IMEMode](Access.ListBox.IMEMode.md)
- [IMESentenceMode](Access.ListBox.IMESentenceMode.md)
- [InheritValueList](Access.ListBox.InheritValueList.md)
- [InSelection](Access.ListBox.InSelection.md)
- [IsVisible](Access.ListBox.IsVisible.md)
- [ItemData](Access.ListBox.ItemData.md)
- [ItemsSelected](Access.ListBox.ItemsSelected.md)
- [LabelAlign](Access.ListBox.LabelAlign.md)
- [LabelX](Access.ListBox.LabelX.md)
- [LabelY](Access.ListBox.LabelY.md)
- [Layout](Access.ListBox.Layout.md)
- [LayoutID](Access.ListBox.LayoutID.md)
- [Left](Access.ListBox.Left.md)
- [LeftPadding](Access.ListBox.LeftPadding.md)
- [ListCount](Access.ListBox.ListCount.md)
- [ListIndex](Access.ListBox.ListIndex.md)
- [ListItemsEditForm](Access.ListBox.ListItemsEditForm.md)
- [Locked](Access.ListBox.Locked.md)
- [MultiSelect](Access.ListBox.MultiSelect.md)
- [Name](Access.ListBox.Name.md)
- [NumeralShapes](Access.ListBox.NumeralShapes.md)
- [OldBorderStyle](Access.ListBox.OldBorderStyle.md)
- [OldValue](Access.ListBox.OldValue.md)
- [OnClick](Access.ListBox.OnClick.md)
- [OnDblClick](Access.ListBox.OnDblClick.md)
- [OnEnter](Access.ListBox.OnEnter.md)
- [OnExit](Access.ListBox.OnExit.md)
- [OnGotFocus](Access.ListBox.OnGotFocus.md)
- [OnKeyDown](Access.ListBox.OnKeyDown.md)
- [OnKeyPress](Access.ListBox.OnKeyPress.md)
- [OnKeyUp](Access.ListBox.OnKeyUp.md)
- [OnLostFocus](Access.ListBox.OnLostFocus.md)
- [OnMouseDown](Access.ListBox.OnMouseDown.md)
- [OnMouseMove](Access.ListBox.OnMouseMove.md)
- [OnMouseUp](Access.ListBox.OnMouseUp.md)
- [Parent](Access.ListBox.Parent.md)
- [Properties](Access.ListBox.Properties.md)
- [ReadingOrder](Access.ListBox.ReadingOrder.md)
- [Recordset](Access.ListBox.Recordset.md)
- [RightPadding](Access.ListBox.RightPadding.md)
- [RowSource](Access.ListBox.RowSource.md)
- [RowSourceType](Access.ListBox.RowSourceType.md)
- [ScrollBarAlign](Access.ListBox.ScrollBarAlign.md)
- [Section](Access.ListBox.Section.md)
- [Selected](Access.ListBox.Selected.md)
- [ShortcutMenuBar](Access.ListBox.ShortcutMenuBar.md)
- [ShowOnlyRowSourceValues](Access.ListBox.ShowOnlyRowSourceValues.md)
- [SmartTags](Access.ListBox.SmartTags.md)
- [SpecialEffect](Access.ListBox.SpecialEffect.md)
- [StatusBarText](Access.ListBox.StatusBarText.md)
- [TabIndex](Access.ListBox.TabIndex.md)
- [TabStop](Access.ListBox.TabStop.md)
- [Tag](Access.ListBox.Tag.md)
- [ThemeFontIndex](Access.ListBox.ThemeFontIndex.md)
- [Top](Access.ListBox.Top.md)
- [TopPadding](Access.ListBox.TopPadding.md)
- [ValidationRule](Access.ListBox.ValidationRule.md)
- [ValidationText](Access.ListBox.ValidationText.md)
- [Value](Access.ListBox.Value.md)
- [VerticalAnchor](Access.ListBox.VerticalAnchor.md)
- [Visible](Access.ListBox.Visible.md)
- [Width](Access.ListBox.Width.md)

## See also

- [Access object model reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
