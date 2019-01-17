---
title: Enabled, EnterFieldBehavior, SelLength, SelStart, SelText properties example
keywords: fm20.chm5225191
f1_keywords:
- fm20.chm5225191
ms.prod: office
ms.assetid: 3a21ec28-9d7e-1b11-9eb9-58907020ba79
ms.date: 11/14/2018
localization_priority: Normal
---


# Enabled, EnterFieldBehavior, SelLength, SelStart, SelText properties example

The following example tracks the selection-related properties (**[SelLength](sellength-property.md)**, **[SelStart](selstart-property.md)**, and **[SelText](seltext-property.md)**) that change as the user moves the insertion point and extends the selection using the keyboard. 

This example also uses the **[Enabled](enabled-property-microsoft-forms.md)** and **[EnterFieldBehavior](enterfieldbehavior-property.md)** properties.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- One large **[TextBox](textbox-control.md)** named TextBox1.    
- Three **TextBox** controls in a column named TextBox2 through TextBox4.
    

```vb
Private Sub TextBox1_KeyUp(ByVal KeyCode As _ 
 MSForms.ReturnInteger, ByVal Shift As Integer) 
 TextBox2.Text = TextBox1.SelStart 
 TextBox3.Text = TextBox1.SelLength 
 TextBox4.Text = TextBox1.SelText 
End Sub
```

<br/>


```vb
Private Sub UserForm_Initialize() 
 TextBox1.MultiLine = True 
 TextBox1.EnterFieldBehavior = _ 
 fmEnterFieldBehaviorRecallSelection 
 
 TextBox1.Text = "Type your text here. Use " _ 
 & "CTRL+ENTER to start a new line." 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]