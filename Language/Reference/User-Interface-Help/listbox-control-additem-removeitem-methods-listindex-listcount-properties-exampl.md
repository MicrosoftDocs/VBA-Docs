---
title: ListBox control, AddItem, RemoveItem methods, ListIndex, ListCount properties example
keywords: fm20.chm5225178
f1_keywords:
- fm20.chm5225178
ms.prod: office
ms.assetid: 70bc2f0c-79a5-89f2-e987-84f673d4bf97
ms.date: 11/14/2018
localization_priority: Normal
---


# ListBox control, AddItem, RemoveItem methods, ListIndex, ListCount properties example

The following example adds and deletes the contents of a **[ListBox](listbox-control.md)** using the **[AddItem](additem-method.md)** and **[RemoveItem](removeitem-method.md)** methods, and the **[ListIndex](listindex-property.md)** and **[ListCount](listcount-property.md)** properties.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **ListBox** named ListBox1.   
- Two **[CommandButton](commandbutton-control.md)** controls named CommandButton1 and CommandButton2.
    

```vb
Dim EntryCount As Single 
 
Private Sub CommandButton1_Click() 
 EntryCount = EntryCount + 1 
 ListBox1.AddItem (EntryCount & " - Selection") 
End Sub
```

<br/>

```vb
Private Sub CommandButton2_Click() 
 'Ensure ListBox contains list items 
 If ListBox1.ListCount >= 1 Then 
 'If no selection, choose last list item. 
 If ListBox1.ListIndex = -1 Then 
 ListBox1.ListIndex = _ 
 ListBox1.ListCount - 1 
 End If 
 ListBox1.RemoveItem (ListBox1.ListIndex) 
 End If 
End Sub
```

<br/>

```vb
Private Sub UserForm_Initialize() 
 EntryCount = 0 
 CommandButton1.Caption = "Add Item" 
 CommandButton2.Caption = "Remove Item" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
