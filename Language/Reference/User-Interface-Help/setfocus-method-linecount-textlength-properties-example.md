---
title: SetFocus method, LineCount, TextLength properties example
keywords: fm20.chm5225149
f1_keywords:
- fm20.chm5225149
ms.prod: office
ms.assetid: 00b01a7c-f5f5-bc90-06a3-7f7a5bb71dc4
ms.date: 11/14/2018
localization_priority: Normal
---


# SetFocus method, LineCount, TextLength properties example

The following example counts the characters and the number of lines of text in a **[TextBox](textbox-control.md)** by using the **[LineCount](linecount-property.md)** and **[TextLength](textlength-property.md)** properties, and the **[SetFocus](setfocus-method.md)** method. 

In this example, the user can type into a **TextBox**, and can retrieve current values of the **LineCount** and **TextLength** properties.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains the following controls:

- A **TextBox** named TextBox1.   
- A **[CommandButton](commandbutton-control.md)** named CommandButton1.    
- Two **[Label](label-control.md)** controls named Label1 and Label2.
    

```vb
'Type SHIFT+ENTER to start a new line in the text box. 
 
Private Sub CommandButton1_Click() 
 'Must first give TextBox1 the focus to get line 
 'count 
 TextBox1.SetFocus 
 Label1.Caption = "LineCount = " _ 
 & TextBox1.LineCount 
 Label2.Caption = "TextLength = " _ 
 & TextBox1.TextLength 
End Sub 
 
Private Sub UserForm_Initialize() 
 CommandButton1.WordWrap = True 
 CommandButton1.AutoSize = True 
 CommandButton1.Caption = "Get Counts" 
 
 Label1.Caption = "LineCount = " 
 Label2.Caption = "TextLength = " 
 
 TextBox1.MultiLine = True 
 TextBox1.WordWrap = True 
 TextBox1.Text = "Enter your text here." 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]