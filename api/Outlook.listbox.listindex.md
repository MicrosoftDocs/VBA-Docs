---
title: ListBox.ListIndex Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: c3eb93ea-bc47-6c2c-f80d-c9b53f797ef3
ms.date: 06/08/2017
localization_priority: Normal
---


# ListBox.ListIndex Property (Outlook Forms Script)

Returns or sets a  **Variant** that represents the currently selected item in a **[ListBox](Outlook.listbox.md)**. Read/write.


## Syntax

_expression_.**ListIndex**

_expression_ A variable that represents a  **ListBox** object.


## Remarks

The  **ListIndex** property contains an index of the selected row in a list. Values of **ListIndex** range from -1 to one less than the total number of rows in a list (that is, **[ListCount](Outlook.listbox.listcount.md)** - 1). When no rows are selected, **ListIndex** returns -1. When the user selects a row in a **ListBox** or **ComboBox**, the system sets the  **ListIndex** value. The **ListIndex** value of the first row in a list is 0, the value of the second row is 1, and so on.

If you use the  **[MultiSelect](Outlook.listbox.multiselect.md)** property to create a **ListBox** that allows multiple selections, the **[Selected](Outlook.listbox.selected.md)** property of the **ListBox** (rather than the **ListIndex** property) identifies the selected rows. The **Selected** property is an array with the same number of values as the number of rows in the **ListBox**. For each row in the list box,  **Selected** is **True** if the row is selected and **False** if it is not. In a **ListBox** that allows multiple selections, **ListIndex** returns the index of the row that has focus, regardless of whether that row is currently selected.

The  **ListIndex** value is also available by setting the **[BoundColumn](Outlook.listbox.boundcolumn.md)** property to 0 for a list box. If **BoundColumn** is 0, the underlying data source to which the list box is bound contains the same list index value as **ListIndex**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]