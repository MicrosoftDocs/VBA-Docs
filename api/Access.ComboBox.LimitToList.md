---
title: ComboBox.LimitToList property (Access)
keywords: vbaac10.chm11387
f1_keywords:
- vbaac10.chm11387
ms.prod: access
api_name:
- Access.ComboBox.LimitToList
ms.assetid: 885ed814-6e04-b9f1-0acb-3ded28e00f93
ms.date: 03/02/2019
localization_priority: Normal
---


# ComboBox.LimitToList property (Access)

You can use the **LimitToList** property to limit a combo box's values to the listed items. Read/write **Boolean**.


## Syntax

_expression_.**LimitToList**

_expression_ A variable that represents a **[ComboBox](Access.ComboBox.md)** object.


## Remarks

The **LimitToList** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|Yes|**True**|If the user selects an item from the list in the combo box or enters text that matches a listed item, Microsoft Access accepts it. If the entered text doesn't match a listed item, the text isn't accepted and the user must then retype the entry, select a listed item, press Esc, or choose **Undo** on the **Edit** menu.|
|No|**False**|(Default) Access accepts any text that conforms to the **[ValidationRule](Access.ComboBox.ValidationRule.md)** property.|

For [table fields](overview/table-field.md), you can set this property on the **Lookup** tab of the Field Properties section of table Design view for fields with the **DisplayControl** property set to Combo Box.

> [!NOTE] 
> Access sets the **LimitToList** property automatically when you select Lookup Wizard as the data type for a field in table Design view.

When the **LimitToList** property of a bound combo box is set to No, you can enter a value in the combo box that isn't included in the list. Access stores the new value in the form's underlying table or query (in the field specified in the combo box's **[ControlSource](Access.ComboBox.ControlSource.md)** property), not the table or query set for the combo box by the **[RowSource](Access.ComboBox.RowSource.md)** property. To have newly entered values appear in the combo box, you must add the new value to the table or query set in the **RowSource** property by using a macro or Visual Basic event procedure that runs when the **NotInList** event occurs.

Setting both the **LimitToList** property and the **[AutoExpand](Access.ComboBox.AutoExpand.md)** property to Yes lets Access find matching values from the list as the user enters characters in the text box portion of the combo box, and restricts the entries to only those values.

> [!NOTE] 
> If you set the combo box's **[BoundColumn](Access.ComboBox.BoundColumn.md)** property to any column other than the first visible column (or if you set **BoundColumn** to 0), the **LimitToList** property is automatically set to Yes.

When the **LimitToList** property is set to Yes and the user selects the arrow next to the combo box, Access selects matching values in the list as the user enters characters in the text box portion of the combo box, even if the **AutoExpand** property is set to No. If the user presses Enter or moves to another control or record, the selected value appears in the combo box.

Combo boxes accept **null** values when the **LimitToList** property is set to Yes or **True**, whether or not the list contains **null** values. If you want to prevent users from entering a **null** value in a combo box, set the **Required** property of the field in the table to which the combo box is bound to Yes.


## Example

The following example limits a given combo box's values to its listed items.

```vb
Forms("Order Entry").Controls("States").LimitToList = True  

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
