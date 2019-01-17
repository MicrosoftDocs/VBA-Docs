---
title: DblClick event, CanPaste property, Paste, Copy methods example
keywords: fm20.chm5225167
f1_keywords:
- fm20.chm5225167
ms.prod: office
ms.assetid: 318cfadf-5e97-0a42-5491-0dbbe077efd4
ms.date: 11/14/2018
localization_priority: Normal
---


# DblClick event, CanPaste property, Paste, Copy methods example

The following example uses the **[CanPaste](canpaste-property.md)** property and the **[Paste](paste-method-microsoft-forms.md)** method to paste a **[ComboBox](combobox-control.md)** from the Clipboard to a **[Page](page-object.md)** of a **[MultiPage](multipage-control.md)**. 

This sample also uses the **[SetFocus](setfocus-method.md)** and **[Copy](copy-method-microsoft-forms.md)** methods to copy a control from the form to the Clipboard.

The user clicks CommandButton1 to copy the **ComboBox** to the Clipboard. The user double-clicks (using the DblClick event) CommandButton1 to paste the **ComboBox** to the **MultiPage**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **[TextBox](textbox-control.md)** named TextBox1.   
- A **ComboBox** named ComboBox1.    
- A **MultiPage** named MultiPage1.    
- A **[CommandButton](commandbutton-control.md)** named CommandButton1.
    

> [!NOTE] 
> This example also includes a subroutine to illustrate pasting text into a control.


```vb
Private Sub UserForm_Initialize() 
 ComboBox1.AddItem "It's a beautiful day!" 
 
 CommandButton1.Caption = "Copy ComboBox to " _ 
 & "Clipboard" 
 CommandButton1.AutoSize = True 
End Sub 
 
Private Sub MultiPage1_DblClick(ByVal Index As Long, _ 
 ByVal Cancel As MSForms.ReturnBoolean) 
 If MultiPage1.Pages(MultiPage1.Value).CanPaste = _ 
 True 
 Then 
 MultiPage1.Pages(MultiPage1.Value).Paste 
 Else 
 TextBox1.Text = "Can't Paste" 
 End If 
End Sub 
 
Private Sub CommandButton1_Click() 
 UserForm1.ComboBox1.SetFocus 
 UserForm1.Copy 
End Sub 
 
'Code for pasting text into a control 
'Private Sub ComboBox1_DblClick(ByVal Cancel As _ 
 MSForms.ReturnBoolean) 
' If ComboBox1.CanPaste = True Then 
' ComboBox1.Paste 
' Else 
' TextBox1.Text = "Can't Paste" 
' End If 
'End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]