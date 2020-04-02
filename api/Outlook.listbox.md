---
title: ListBox object (Outlook Forms Script)
keywords: olfm10.chm2000560
f1_keywords:
- olfm10.chm2000560
ms.prod: outlook
ms.assetid: f56ba480-f8fe-6d12-265e-3b0a9838af97
ms.date: 06/08/2017
localization_priority: Normal
---


# ListBox object (Outlook Forms Script)

Displays a list of values and lets you select one or more.


## Remarks

If the **ListBox** is bound to a data source, the **ListBox** stores the selected value in that data source.

The **ListBox** can either appear as a list or as a group of **[OptionButton](Outlook.optionbutton.md)** controls or **[CheckBox](Outlook.checkbox.md)** controls.

The default property for a **ListBox** is the **[Value](Outlook.listbox.value.md)** property.

The default event for a **ListBox** is the **[Click](Outlook.listbox.click.md)** event.

You can't drop text into a drop-down **ListBox**.


### ListBox styles

You can choose between two presentation styles for a **ListBox**. This is expressed by the **[ListStyle](Outlook.listbox.liststyle.md)** property. Each style provides different ways for users to select items in the list.

If the style is 0, each item is on a separate row; the user selects an item by highlighting one or more rows.

If the style is 1, an **OptionButton** or **CheckBox** appears at the beginning of each row. With this style, the user selects an item by clicking the option button or check box. Check boxes appear only when the **[MultiSelect](Outlook.listbox.multiselect.md)** property is **True**.


## Events

|Name|Description|
|:-----|:-----|
| [Click](Outlook.ListBox.click.md)|Occurs when the user definitively selects a value for the control that has more than one possible value.|


## Methods

|Name|Description|
|:-----|:-----|
| [AddItem](Outlook.ListBox.additem.md)|For a single-column [ListBox](Outlook.ListBox.md), the **AddItem** method adds an item to the list. For a multicolumn **ListBox**, this method adds a row to the list.|
| [Clear](Outlook.ListBox.clear.md)|Removes all entries in the list in a **ListBox**.|
| [RemoveItem](Outlook.ListBox.removeitem.md)|Removes a row from the list in a **ListBox**.|


## Properties

|Name|Description|
|:-----|:-----|
| [BackColor](Outlook.ListBox.backcolor.md)|Returns or sets a **Long** that specifies the background color of the object. Read/write.|
| [BorderColor](Outlook.ListBox.bordercolor.md)|Returns or sets a **Long** that specifies the border color of an object. Read/write.|
| [BorderStyle](Outlook.ListBox.borderstyle.md)|Returns or sets an **Integer** that specifies the type of border of the control. Read/write.|
| [BoundColumn](Outlook.ListBox.boundcolumn.md)|Returns or sets a **Variant** that identifies the source of data in a multicolumn **ListBox**. Read/write.|
| [Column](Outlook.ListBox.column.md)|Returns or sets a **Variant** that represents a single value, a column of values, or a two-dimensional array to load into a **ListBox**. Read/write.|
| [ColumnCount](Outlook.ListBox.columncount.md)|Returns or sets a **Long** that represents the number of columns to display in a list box. Read/write.|
| [ColumnHeads](Outlook.ListBox.columnheads.md)|Returns or sets a **Boolean** that specifies whether a single row of column headings are displayed. Read/write.|
| [ColumnWidths](Outlook.ListBox.columnwidths.md)|Returns or sets a **String** that specifies the width of each column in a multicolumn **ListBox**. Read/write.|
| [Enabled](Outlook.ListBox.enabled.md)|Returns or sets a **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.|
| [ForeColor](Outlook.ListBox.forecolor.md)|Returns or sets a **Long** that specifies the foreground color of an object. Read/write.|
| [IMEMode](Outlook.ListBox.imemode.md)|Returns or sets an **Integer** that specifies the default run-time mode of the Input Method Editor (IME) for a control. Read/write.|
| [IntegralHeight](Outlook.ListBox.integralheight.md)|Returns or sets a **Boolean** that specifies whether a **ListBox** displays full lines of text in a list or partial lines. Read/write.|
| [List](Outlook.ListBox.list.md)|Returns or sets a **Variant** that represents the specified entry in a **ListBox**. Read/write.|
| [ListCount](Outlook.ListBox.listcount.md)|Returns a **Long** that represents the number of list entries in a control. Read-only.|
| [ListIndex](Outlook.ListBox.listindex.md)|Returns or sets a **Variant** that represents the currently selected item in a **ListBox**. Read/write.|
| [ListStyle](Outlook.ListBox.liststyle.md)|Returns or sets an **Integer** that specifies the visual appearance of the list in a **ListBox**. Read/write.|
| [Locked](Outlook.ListBox.locked.md)|Returns or sets a **Boolean** that specifies whether a control can be edited. Read/write.|
| [MatchEntry](Outlook.ListBox.matchentry.md)|Returns or sets an **Integer** that indicates how a **ListBox** searches its list as the user types. Read/write.|
| [MouseIcon](Outlook.ListBox.mouseicon.md)|Returns a **String** that represents the full path name of a custom icon that is to be assigned to the control. Read-only.|
| [MousePointer](Outlook.ListBox.mousepointer.md)|Returns or sets an **Integer** that specifies the type of pointer displayed when the user positions the mouse over a particular object. Read/write.|
| [MultiSelect](Outlook.ListBox.multiselect.md)|Returns or sets an **Integer** that indicates whether the object permits multiple selections. Read/write.|
| [Selected](Outlook.ListBox.selected.md)|Returns or sets a **Boolean** that indicates the selection state of items in a **ListBox**. Read/write.|
| [SpecialEffect](Outlook.ListBox.specialeffect.md)|Returns or sets an **Integer** that specifies the visual appearance of an object. Read/write.|
| [Text](Outlook.ListBox.text.md)|Returns or sets a **String** that specifies text in a **ListBox**, changing the selected row in the control. Read/write.|
| [TextAlign](Outlook.ListBox.textalign.md)|Returns or sets an **Integer** that specifies how text is aligned in a control. Read/write.|
| [TextColumn](Outlook.ListBox.textcolumn.md)|Returns or sets a **Variant** that identifies the column in a **ListBox** to display to the user. Read/write.|
| [TopIndex](Outlook.ListBox.topindex.md)|Returns or sets a **Long** that represents the index of the list item displayed in the topmost position in the list. Read/write.|
| [Value](Outlook.ListBox.value.md)|Returns or sets a **Variant** that specifies the value in the [BoundColumn](Outlook.ListBox.boundcolumn.md) of the currently selected rows. Read/write.|





[!include[Support and feedback](~/includes/feedback-boilerplate.md)]