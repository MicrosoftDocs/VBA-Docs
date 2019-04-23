---
title: Paste, PutInClipboard, SetText methods example
keywords: fm20.chm5225164
f1_keywords:
- fm20.chm5225164
ms.prod: office
ms.assetid: d7045eb8-3b79-a490-91a8-b6f5369bbf8c
ms.date: 11/14/2018
localization_priority: Normal
---


# Paste, PutInClipboard, SetText methods example

The following example demonstrates data movement from a **[TextBox](textbox-control.md)** to a **[DataObject](dataobject-object.md)**, from a **[DataObject](dataobject-object.md)** to the Clipboard, and from the Clipboard to another **TextBox**. 

The **[PutInClipboard](putinclipboard-method.md)** method transfers the data from a **DataObject** to the Clipboard. The **[SetText](settext-method.md)** and **[Paste](paste-method-microsoft-forms.md)** methods are also used.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- Two **TextBox** controls named TextBox1 and TextBox2.  
- A **[CommandButton](commandbutton-control.md)** named CommandButton1.
    
```vb
Dim MyData As DataObject 
 
Private Sub CommandButton1_Click() 
 Set MyData = New DataObject 
 
 MyData.SetText TextBox1.Text 
 MyData.PutInClipboard 
 
 TextBox2.Paste 
End Sub 
 
Private Sub UserForm_Initialize() 
 TextBox1.Text = "Move this data to a " _ 
 & "DataObject, to the Clipboard, then to " _ 
 & "TextBox2!" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
