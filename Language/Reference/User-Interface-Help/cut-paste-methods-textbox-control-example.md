---
title: Cut, Paste methods, TextBox control example
keywords: fm20.chm5225154
f1_keywords:
- fm20.chm5225154
ms.prod: office
ms.assetid: 38f39c6b-ff99-a5ca-596a-e2ddace29324
ms.date: 11/14/2018
localization_priority: Normal
---


# Cut, Paste methods, TextBox control example

The following example uses the **[Cut](cut-method-microsoft-forms.md)** and **[Paste](paste-method-microsoft-forms.md)** methods to cut text from one **[TextBox](textbox-control.md)** and paste it into another **TextBox**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- Two **TextBox** controls named TextBox1 and TextBox2.    
- A **[CommandButton](commandbutton-control.md)** named CommandButton1.
    

```vb
Private Sub UserForm_Initialize() 
 TextBox1.Text = "From TextBox1!" 
 TextBox2.Text = "Hello " 
 
 CommandButton1.Caption = "Cut and Paste" 
 CommandButton1.AutoSize = True 
End Sub 
 
Private Sub CommandButton1_Click() 
 TextBox2.SelStart = 0 
 TextBox2.SelLength = TextBox2.TextLength 
 TextBox2.Cut 
 
 TextBox1.SetFocus 
 TextBox1.SelStart = 0 
 
 TextBox1.Paste 
 TextBox2.SelStart = 0 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]