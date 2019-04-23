---
title: Add, Cut, Paste methods, Page object, MultiPage control example
keywords: fm20.chm5225155
f1_keywords:
- fm20.chm5225155
ms.prod: office
ms.assetid: 938475c8-b6cb-88b0-379d-398f52e5c51d
ms.date: 11/14/2018
localization_priority: Normal
---


# Add, Cut, Paste methods, Page object, MultiPage control example

The following example uses the **[Add](add-method-microsoft-forms.md)**, **[Cut](cut-method-microsoft-forms.md)**, and **[Paste](paste-method-microsoft-forms.md)** methods to cut and paste a control from a **[Page](page-object.md)** of a **[MultiPage](multipage-control.md)**. The control involved in the cut and paste operations is dynamically added to the form.

This example assumes that the user will add, cut, and then paste the new control.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- Three **[CommandButton](commandbutton-control.md)** controls named CommandButton1 through CommandButton3.   
- A **MultiPage** named MultiPage1.

```vb
Dim MyTextBox As Control 
 
Private Sub CommandButton1_Click() 
 Set MyTextBox = MultiPage1.Pages(MultiPage1.Value).Controls_ 
 .Add("MSForms.TextBox.1", "MyTextBox", Visible) 
 CommandButton2.Enabled = True 
 CommandButton1.Enabled = False 
End Sub 
 
Private Sub CommandButton2_Click() 
 MultiPage1.Pages(MultiPage1.Value).Controls.Cut 
 CommandButton3.Enabled = True 
 CommandButton2.Enabled = False 
End Sub 
 
Private Sub CommandButton3_Click() 
 Dim MyPage As Object 
 Set MyPage = _ 
 MultiPage1.Pages.Item(MultiPage1.Value) 
 
 MyPage.Paste 
 CommandButton3.Enabled = False 
End Sub 
 
Private Sub UserForm_Initialize() 
 CommandButton1.Caption = "Add" 
 CommandButton2.Caption = "Cut" 
 CommandButton3.Caption = "Paste" 
 
 CommandButton1.Enabled = True 
 CommandButton2.Enabled = False 
 CommandButton3.Enabled = False 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]