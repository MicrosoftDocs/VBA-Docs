---
title: Page object, MultiPage control, Add, Clear, Remove methods example
keywords: fm20.chm5225177
f1_keywords:
- fm20.chm5225177
ms.prod: office
ms.assetid: ba40e297-6f1f-b012-34a2-d8e6c6b0e462
ms.date: 11/14/2018
localization_priority: Normal
---


# Page object, MultiPage control, Add, Clear, Remove methods example

The following example uses the **[Add](add-method-microsoft-forms.md)**, **[Clear](clear-method-microsoft-forms.md)**, and **[Remove](remove-method.md)** methods to add and remove a control to a **[Page](page-object.md)** of a **[MultiPage](multipage-control.md)** at run time.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **MultiPage** named MultiPage1.   
- Three **[CommandButton](commandbutton-control.md)** controls named CommandButton1 through CommandButton3.
    

```vb
Dim MyTextBox As Control 
 
Private Sub CommandButton1_Click() 
Set MyTextBox = MultiPage1.Pages(0).Controls.Add("MSForms" _ 
 & ".TextBox.1", "MyTextBox", Visible) 
End Sub 
 
Private Sub CommandButton2_Click() 
 MultiPage1.Pages(0).Controls.Clear 
End Sub 
 
Private Sub CommandButton3_Click() 
 If MultiPage1.Pages(0).Controls.Count > 0 Then 
 MultiPage1.Pages(0).Controls.Remove "MyTextBox" 
 End If 
End Sub 
 
Private Sub UserForm_Initialize() 
 CommandButton1.Caption = "Add control" 
 CommandButton2.Caption = "Clear controls" 
 CommandButton3.Caption = "Remove control" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]