---
title: TopIndex Property Example
keywords: fm20.chm5225130
f1_keywords:
- fm20.chm5225130
ms.prod: office
ms.assetid: 6b88e7dd-1b2f-0b1a-2348-986bf97461c9
ms.date: 06/08/2017
---


# TopIndex Property Example

The following example identifies the top item displayed in a  **[ListBox](listbox-control.md)** and the item that has the focus within the **[ListBox](listbox-control.md)**. This example uses the **TopIndex** property to identify the item displayed at the top of the **[ListBox](listbox-control.md)** and the **ListIndex** property to identify the item that has the focus. The user selects an item in the **[ListBox](listbox-control.md)**. The displayed values of **TopIndex** and **ListIndex** are updated when the user selects an item or when the user clicks the **[CommandButton](commandbutton-control.md)**..

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **[Label](label-control.md)** named Label1.
    
- A  **[TextBox](textbox-control.md)** named TextBox1.
    
- A  **[Label](label-control.md)** named Label2.
    
- A  **[TextBox](textbox-control.md)** named TextBox2.
    
- A  **[CommandButton](commandbutton-control.md)** named CommandButton1.
    
- A  **[ListBox](listbox-control.md)** named ListBox1.
    




```vb
Private Sub CommandButton1_Click() 
 ListBox1.TopIndex = ListBox1.ListIndex 
 TextBox1.Text = ListBox1.TopIndex 
 TextBox2.Text = ListBox1.ListIndex 
End Sub 
 
Private Sub ListBox1_Change() 
 TextBox1.Text = ListBox1.TopIndex 
 TextBox2.Text = ListBox1.ListIndex 
End Sub 
 
Private Sub UserForm_Initialize() 
 Dim i As Integer 
 
 For i = 0 To 24 
 ListBox1.AddItem "Choice " & (i + 1) 
 Next i 
 ListBox1.Height = 66 
 CommandButton1.Caption = "Move to top of list" 
 CommandButton1.AutoSize = True 
 CommandButton1.TakeFocusOnClick = False 
 
 Label1.Caption = "Index of top item" 
 TextBox1.Text = ListBox1.TopIndex 
 
 Label2. Caption = "Index of current item" 
 Label2.AutoSize = True 
 Label2.WordWrap = False 
 TextBox2.Text = ListBox1.ListIndex 
End Sub
```


