---
title: Copy, GetFromClipboard, GetText methods, DataObject object example
keywords: fm20.chm5225163
f1_keywords:
- fm20.chm5225163
ms.prod: office
ms.assetid: 6be27a7e-58f2-5cad-5ed0-570520fd61f1
ms.date: 11/14/2018 
localization_priority: Normal
---


# Copy, GetFromClipboard, GetText methods, DataObject object example

The following example demonstrates data movement from a **[TextBox](textbox-control.md)** to the Clipboard, from the Clipboard to a **[DataObject](dataobject-object.md)**, and from a **DataObject** into another **TextBox**. The **[GetFromClipboard](getfromclipboard-method.md)** method transfers the data from the Clipboard to a **DataObject**. The **[Copy](copy-method-microsoft-forms.md)** and **[GetText](gettext-method-microsoft-forms.md)** methods are also used.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- Two **TextBox** controls named TextBox1 and TextBox2.   
- A **[CommandButton](commandbutton-control.md)** named CommandButton1.
    

```vb
Dim MyData as DataObject 
 
Private Sub CommandButton1_Click() 
 'Need to select text before copying it to Clipboard 
 TextBox1.SelStart = 0 
 TextBox1.SelLength = TextBox1.TextLength 
 TextBox1.Copy 
 
 MyData.GetFromClipboard 
 TextBox2.Text = MyData.GetText(1) 
End Sub 
 
Private Sub UserForm_Initialize() 
 Set MyData = New DataObject 
 TextBox1.Text = "Move this data to the " _ 
 & "Clipboard, to a DataObject, then to " 
 & "TextBox2!" 
End Sub 
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
