---
title: KeyDown event, TextBox control, CurLine, CurTargetX, CurX, Text properties example
keywords: fm20.chm5225187
f1_keywords:
- fm20.chm5225187
ms.prod: office
ms.assetid: 696c6429-7a62-9eeb-d7c3-a883e888da09
ms.date: 11/14/2018
localization_priority: Normal
---


# KeyDown event, TextBox control, CurLine, CurTargetX, CurX, Text properties example

The following example tracks the **[CurLine](curline-property.md)**, **[CurTargetX](curtargetx-property.md)**, and **[CurX](curx-property.md)** property settings in a multiline **[TextBox](textbox-control.md)**. These settings change in the **[KeyUp](keydown-keyup-events.md)** event as the user types into the **[Text](text-property-microsoft-forms.md)** property, moves the insertion point, and extends the selection by using the keyboard.

To use this example, follow these steps:

1. Copy this sample code to the Declarations portion of a form.
    
2. Add one large **TextBox** named TextBox1 to the form.
    
3. Add three **TextBox** controls named TextBox2, TextBox3, and TextBox4 in a column.
    

```vb
Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) 
 TextBox2.Text = TextBox1.CurLine 
 TextBox3.Text = TextBox1.CurX 
 TextBox4.Text = TextBox1.CurTargetX 
End Sub
```

<br/>


```vb
Private Sub UserForm_Initialize() 
 TextBox1.MultiLine = True 
 
 TextBox1.Text = "Type your text here. User CTRL + ENTER to start a new line." 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]