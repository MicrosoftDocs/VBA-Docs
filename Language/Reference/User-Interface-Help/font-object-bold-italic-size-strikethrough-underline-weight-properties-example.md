---
title: Font object, Bold, Italic, Size, StrikeThrough, Underline, Weight properties example
keywords: fm20.chm5225182
f1_keywords:
- fm20.chm5225182
ms.prod: office
ms.assetid: b7fc7c3e-b7ef-9ff3-1dde-06792abf4c51
ms.date: 11/14/2018
localization_priority: Normal
---


# Font object, Bold, Italic, Size, StrikeThrough, Underline, Weight properties example

The following example demonstrates a **[Font](font-object-microsoft-forms.md)** object and the **[Bold, Italic, Size, StrikeThrough, Underline, Weight](bold-italic-size-strikethrough-underline-weight-properties.md)** properties related to fonts. You can manipulate font properties of an object directly or by using an alias, as this example also shows.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **[Label](label-control.md)** named Label1.   
- Four **[ToggleButton](togglebutton-control.md)** controls named ToggleButton1 through ToggleButton4.    
- A second **Label** and a **[TextBox](textbox-control.md)** named Label2 and TextBox1.
    

```vb
Dim MyFont As StdFont 
 
Private Sub ToggleButton1_Click() 
 If ToggleButton1.Value = True Then 
 MyFont.Bold = True 
 'Using MyFont alias to control font 
 ToggleButton1.Caption = "Bold On" 
 MyFont.Size = 22 
 'Increase the font size 
 Else 
 MyFont.Bold = False 
 ToggleButton1.Caption = "Bold Off" 
 MyFont.Size = 8 
 'Return font size to initial size 
 End If 
 
 TextBox1.Text = Str(MyFont.Weight) 
 'Bold and Weight are related 
End Sub 
 
Private Sub ToggleButton2_Click() 
 If ToggleButton2.Value = True Then 
 Label1.Font.Italic = True 
 'Using Label1.Font directly 
 ToggleButton2.Caption = "Italic On" 
 Else 
 Label1.Font.Italic = False 
 ToggleButton2.Caption = "Italic Off" 
 End If 
End Sub 
 
Private Sub ToggleButton3_Click() 
 If ToggleButton3.Value = True Then 
 Label1.Font.Strikethrough = True 
 'Using Label1.Font directly 
 ToggleButton3.Caption = "StrikeThrough On" 
 Else 
 Label1.Font.Strikethrough = False 
 ToggleButton3.Caption = "StrikeThrough Off" 
 End If 
End Sub 
 
Private Sub ToggleButton4_Click() 
 If ToggleButton4.Value = True Then 
 MyFont.Underline = True 
 'Using MyFont alias for Label1.Font 
 ToggleButton4.Caption = "Underline On" 
 Else 
 Label1.Font.Underline = False 
 ToggleButton4.Caption = "Underline Off" 
 End If 
End Sub 
 
Private Sub UserForm_Initialize() 
 Set MyFont = Label1.Font 
 
 ToggleButton1.Value = True 
 ToggleButton1.Caption = "Bold On" 
 
 Label1.AutoSize = True 'Set size of Label1 
 Label1.AutoSize = False 
 
 ToggleButton2.Value = False 
 ToggleButton2.Caption = "Italic Off" 
 
 ToggleButton3.Value = False 
 ToggleButton3.Caption = "StrikeThrough Off" 
 
 ToggleButton4.Value = False 
 ToggleButton4.Caption = "Underline Off" 
 
 Label2.Caption = "Font Weight" 
 TextBox1.Text = Str(Label1.Font.Weight) 
 TextBox1.Enabled = False 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
