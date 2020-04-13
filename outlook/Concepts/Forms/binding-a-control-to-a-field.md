---
title: Binding a Control to a Field
ms.prod: outlook
ms.assetid: 8e338547-b3ff-b84b-16b9-0c465256d972
ms.date: 06/08/2019
localization_priority: Normal
---


# Binding a Control to a Field

If you have created a control using the **ontrol Toolbox** and you would like the information in this control to be saved with the form, the control must be bound to a field. This means that a link will be established between the control and the source of data, in this case, a field in the item.

1. In the Forms Designer, right-click the control you want to bind to a field and then click **properties** on the shortcut menu.

2. On the **alue** tab, click **Choose Field**, point to a field set, and then click the field to which you want to bind the control. If you would like to bind the control to a new field that you create, click **w**. In the **Field** box, type the name of your new field in the **Name**: area. Click ** ** in the **New Field** box.

3. Click **K** in the **Properties** box.

 **Note** If you bind a **CheckBox](../../../api/Outlook.checkbox.md)**, **omboBox](../../../api/Outlook.combobox.md)**, * **stBox](../../../api/Outlook.listbox.md)**, or ** **ionButton](../../../api/Outlook.optionbutton.md)** to a field, then the **Click** event does not fire. You need to use the **PropertyChange** or **CustomPropertyChange** event of the item to detect the change via code, as shown in the following example:

```vb
Sub Item_PropertyChange(ByVal Name) 
Set MyListBox = Item.GetInspector.ModifiedFormPages("Message").Controls("ListBox1") 
Select Case Name 
 Case "Mileage" 
 Item.CC = MyListBox.Value 
 Item.Subject = MyListBox.Value 
 Case Else 
End Select 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]