---
title: ComboBox object (Outlook Forms Script)
keywords: olfm10.chm2000480
f1_keywords:
- olfm10.chm2000480
ms.prod: outlook
ms.assetid: 31e7c1de-ee4e-b3d9-4579-7fc6b215bad3
ms.date: 06/08/2017
localization_priority: Normal
---


# ComboBox object (Outlook Forms Script)

Combines the features of a **[ListBox](Outlook.listbox.md)** and a **[TextBox](Outlook.textbox.md)**. 


## Remarks

The user can enter a new value, as with a **TextBox**, or the user can select an existing value as with a **ListBox**.

If a **ComboBox** is bound to a data source, the **ComboBox** inserts the value entered or selected by the user into that data source. If a multicolumn combo box is bound, then the **[BoundColumn](Outlook.combobox.boundcolumn.md)** property determines which value is stored in the bound data source.

The list in a **ComboBox** consists of rows of data. Each row can have one or more columns, which can appear with or without headings. Some applications do not support column headings, others provide only limited support.

The default property of a **ComboBox** is the **[Value](Outlook.combobox.value.md)** property.

If you want more than a single line of the list to appear at all times, you might want to use a **ListBox** instead of a **ComboBox**. If you want to use a **ComboBox** and limit values to those in the list, you can set the **[Style](Outlook.combobox.style.md)** property of the **ComboBox** so the control looks like a drop-down list box.


## Events

|Name|Description|
|:-----|:-----|
| [Click](Outlook.combobox.click.md)|Occurs when the user definitively selects a value for the control that has more than one possible value.|


## Methods

|Name|Description|
|:-----|:-----|
| [AddItem](Outlook.combobox.additem.md)|For a single-column  [ComboBox](Outlook.combobox.md), the **AddItem** method adds an item to the list. For a multicolumn **ComboBox**, this method adds a row to the list.|
| [Clear](Outlook.combobox.clear.md)|Removes all entries in the list in a **ComboBox**.|
| [Copy](Outlook.combobox.copy.md)|Copies the contents of an object to the Clipboard.|
| [Cut](Outlook.combobox.cut.md)|Removes selected information from an object and transfers it to the Clipboard.|
| [DropDown](Outlook.combobox.dropdown.md)|Displays the list portion of a **ComboBox**.|
| [Paste](Outlook.combobox.paste.md)|Transfers the contents of the Clipboard to an object.|
| [RemoveItem](Outlook.combobox.removeitem.md)|Removes a row from the list in a **ComboBox**.|


## Properties

|Name|Description|
|:-----|:-----|
| [AutoSize](Outlook.combobox.autosize.md)|Returns or sets a **Boolean** that specifies whether an object automatically resizes to display its entire contents. Read/write.|
| [AutoTab](Outlook.combobox.autotab.md)|Returns or sets a **Boolean** that specifies whether an automatic tab occurs when a user enters the maximum allowable number of characters into the text box portion of a **ComboBox**. Read/write.|
| [AutoWordSelect](Outlook.combobox.autowordselect.md)|Returns or sets a **Boolean** that specifies whether the basic unit used to extend a selection is a word or a single character. Read/write.|
| [BackColor](Outlook.combobox.backcolor.md)|Returns or sets a **Long** that specifies the background color of the object. Read/write.|
| [BackStyle](Outlook.combobox.backstyle.md)|Returns or sets an **Integer** that specifies the background style for an object. Read/write.|
| [BorderColor](Outlook.combobox.bordercolor.md)|Returns or sets a **Long** that specifies the border color of an object. Read/write.|
| [BorderStyle](Outlook.combobox.borderstyle.md)|Returns or sets an **Integer** that specifies the type of border of the control. Read/write.|
| [BoundColumn](Outlook.combobox.boundcolumn.md)|Returns or sets a **Variant** that identifies the source of data in a multicolumn [ComboBox](Outlook.combobox.md). Read/write.|
| [CanPaste](Outlook.combobox.canpaste.md)|Returns a **Boolean** that specifies whether the Clipboard contains data that the object supports. Read-only.|
| [Column](Outlook.combobox.column.md)|Returns or sets a **Variant** that represents a single value, a column of values, or a two-dimensional array to load into a **ComboBox**. Read/write.|
| [ColumnCount](Outlook.combobox.columncount.md)|Returns or sets a **Long** that represents the number of columns to display in a combo box. Read/write.|
| [ColumnHeads](Outlook.combobox.columnheads.md)|Returns or sets a **Boolean** that specifies whether a single row of column headings are displayed. Read/write.|
| [ColumnWidths](Outlook.combobox.columnwidths.md)|Returns or sets a **String** that specifies the width of each column in a multicolumn **ComboBox**. Read/write.|
| [CurTargetX](Outlook.combobox.curtargetx.md)|Returns a **Long** that represents the preferred horizontal position of the insertion point in a multiline **ComboBox**. Read-only.|
| [CurX](Outlook.combobox.curx.md)|Returns or sets a **Long** that represents the current horizontal position of the insertion point in a multiline **ComboBox**. Read/write.|
| [DragBehavior](Outlook.combobox.dragbehavior.md)|Returns or sets an **Integer** that specifies whether the system enables the drag-and-drop feature for the control. Read/write.|
| [DropButtonStyle](Outlook.combobox.dropbuttonstyle.md)|Returns or sets a **fmDropButtonStyle** value that represents the symbol displayed on the drop button in a **ComboBox**. Read/write.|
| [Enabled](Outlook.combobox.enabled.md)|Returns or sets a **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.|
| [EnterFieldBehavior](Outlook.combobox.enterfieldbehavior.md)|Returns or sets an **Integer** that specifies the selection behavior when entering a **ComboBox**. Read/write.|
| [ForeColor](Outlook.combobox.forecolor.md)|Returns or sets a **Long** that specifies the foreground color of an object. Read/write.|
| [HideSelection](Outlook.combobox.hideselection.md)|Returns or sets a **Boolean** that specifies whether selected text remains highlighted when a control does not have the focus. Read/write.|
| [IMEMode](Outlook.combobox.imemode.md)|Returns or sets an **Integer** that specifies the default run-time mode of the Input Method Editor (IME) for a control. Read/write.|
| [LineCount](Outlook.combobox.linecount.md)|Returns a **Long** that specifies the number of text lines in a **ComboBox**. Read-only.|
| [List](Outlook.combobox.list.md)|Returns or sets a **Variant** that represents the specified entry in a **ComboBox**. Read/write.|
| [ListCount](Outlook.combobox.listcount.md)|Returns a **Long** that represents the number of list entries in a control. Read-only.|
| [ListIndex](Outlook.combobox.listindex.md)|Returns or sets a **Variant** that represents the currently selected item in a **ComboBox**. Read/write.|
| [ListRows](Outlook.combobox.listrows.md)|Returns or sets a **Long** that specifies the maximum number of rows to display in the list. Read/write.|
| [ListStyle](Outlook.combobox.liststyle.md)|Returns or sets an **Integer** that specifies the visual appearance of the list in a **ComboBox**. Read/write.|
| [ListWidth](Outlook.combobox.listwidth.md)|Returns or sets a **Variant** that specifies the width of the list in a **ComboBox**. Read/write.|
| [Locked](Outlook.combobox.locked.md)|Returns or sets a **Boolean** that specifies whether a control can be edited. Read/write.|
| [MatchEntry](Outlook.combobox.matchentry.md)|Returns or sets an **Integer** that indicates how a **ComboBox** searches its list as the user types. Read/write.|
| [MatchFound](Outlook.combobox.matchfound.md)|Returns a **Boolean** value that indicates whether the text that a user has typed into a **ComboBox** matches any of the entries in the list. Read-only.|
| [MatchRequired](Outlook.combobox.matchrequired.md)|Returns or sets a **Boolean** that specifies whether a value entered in the text portion of a **ComboBox** must match an entry in the existing list portion of the control. Read/write.|
| [MaxLength](Outlook.combobox.maxlength.md)|Returns or sets a **Long** that specifies the maximum number of characters a user can enter in a **ComboBox**. Read/write.|
| [MouseIcon](Outlook.combobox.mouseicon.md)|Returns a **String** that represents the full path name of a custom icon that is to be assigned to the control. Read-only.|
| [MousePointer](Outlook.combobox.mousepointer.md)|Returns or sets an **Integer** that specifies the type of pointer displayed when the user positions the mouse over a particular object. Read/write.|
| [SelectionMargin](Outlook.combobox.selectionmargin.md)|Returns or sets a **Boolean** that specifies whether the user can select a line of text by clicking in the region to the left of the text. Read/write.|
| [SelLength](Outlook.combobox.sellength.md)|Returns or sets a **Long** that represents the number of characters selected in the text portion of a **ComboBox**. Read/write.|
| [SelStart](Outlook.combobox.selstart.md)|Returns or sets a **Long** that represents the starting point of selected text, or the insertion point if no text is selected. Read/write.|
| [SelText](Outlook.combobox.seltext.md)|Returns or sets a **String** that represents the selected text of a control. Read/write.|
| [ShowDropButtonWhen](Outlook.combobox.showdropbuttonwhen.md)|Returns or sets a **fmShowDropButtonWhen** value that specifies when to show the drop-down button for a **ComboBox**. Read/write.|
| [SpecialEffect](Outlook.combobox.specialeffect.md)|Returns or sets an **Integer** that specifies the visual appearance of an object. Read/write.|
| [Style](Outlook.combobox.style.md)|Returns or sets an **Integer** that specifies how the user can choose or set the control's value. Read/write.|
| [Text](Outlook.combobox.text.md)|Returns or sets a **String** that specifies text in a **ComboBox**, changing the selected row in the control. Read/write.|
| [TextAlign](Outlook.combobox.textalign.md)|Returns or sets an **Integer** that specifies how text is aligned in a control. Read/write.|
| [TextColumn](Outlook.combobox.textcolumn.md)|Returns or sets a **Variant** that identifies the column in a **ComboBox** to display to the user. Read/write.|
| [TextLength](Outlook.combobox.textlength.md)|Returns a **Long** that represents the length, in number of characters, of text in the edit region of a **ComboBox**. Read-only.|
| [TopIndex](Outlook.combobox.topindex.md)|Returns or sets a **Long** that represents the index of the item displayed in the topmost position in the list portion of the **ComboBox**. Read/write.|
| [Value](Outlook.combobox.value.md)|Returns or sets a **Variant** that specifies the value in the [BoundColumn](Outlook.combobox.boundcolumn.md) of the currently selected rows. Read/write.|




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]