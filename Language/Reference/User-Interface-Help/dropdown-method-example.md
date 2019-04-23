---
title: DropDown method example
keywords: fm20.chm5225153
f1_keywords:
- fm20.chm5225153
ms.prod: office
ms.assetid: 0a450210-9e10-d1f0-cb01-567115c9bfda
ms.date: 11/14/2018
localization_priority: Normal
---


# DropDown method example

The following example uses the **[DropDown](dropdown-method.md)** method to display the list in a **[ComboBox](combobox-control.md)**. The user can display the list of a **ComboBox** by clicking the **[CommandButton](commandbutton-control.md)**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **ComboBox** named ComboBox1.    
- A **CommandButton** named CommandButton1.
    

```vb
Private Sub CommandButton1_Click() 
 ComboBox1.DropDown 
End Sub 
 
Private Sub UserForm_Initialize() 
 ComboBox1.AddItem "Turkey" 
 ComboBox1.AddItem "Chicken" 
 ComboBox1.AddItem "Duck" 
 ComboBox1.AddItem "Goose" 
 ComboBox1.AddItem "Grouse" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]