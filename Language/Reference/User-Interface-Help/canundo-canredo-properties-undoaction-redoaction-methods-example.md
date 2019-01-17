---
title: CanUndo, CanRedo properties, UndoAction, RedoAction methods example
keywords: fm20.chm5225169
f1_keywords:
- fm20.chm5225169
ms.prod: office
ms.assetid: 4c32245c-e209-9343-8351-9fc709b31e66
ms.date: 11/14/2018
localization_priority: Normal
---


# CanUndo, CanRedo properties, UndoAction, RedoAction methods example

The following example demonstrates how to undo or redo text editing within a text box or within the text area of a **[ComboBox](combobox-control.md)**. This sample checks whether an undo or redo operation can occur and then performs the appropriate action. The sample uses the **[CanUndo](canundo-property.md)** and **[CanRedo](canredo-property.md)** properties, and the **[UndoAction](undoaction-method.md)** and **[RedoAction](redoaction-method.md)** methods.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **[TextBox](textbox-control.md)** named TextBox1.   
- A **ComboBox** named ComboBox1.   
- Two **[CommandButton](commandbutton-control.md)** controls named CommandButton1 and CommandButton2.
    

```vb
Private Sub CommandButton1_Click() 
 If UserForm1.CanUndo = True Then 
 UserForm1.UndoAction 
 MsgBox "Undid IT" 
 Else 
 MsgBox "No undo performed." 
 End If 
End Sub 
 
Private Sub CommandButton2_Click() 
 If UserForm1.CanRedo = True Then 
 UserForm1.RedoAction 
 MsgBox "Redid IT" 
 Else 
 MsgBox "No redo performed." 
 End If 
End Sub 
 
Private Sub UserForm_Initialize() 
 TextBox1.Text = "Type your text here." 
 
 ComboBox1.ColumnCount = 3 
 ComboBox1.AddItem "Choice 1, column 1" 
 ComboBox1.List(0, 1) = "Choice 1, column 2" 
 ComboBox1.List(0, 2) = "Choice 1, column 3" 
 
 CommandButton1.Caption = "Undo" 
 CommandButton2.Caption = "Redo" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]