---
title: InsideHeight, InsideWidth properties example
keywords: fm20.chm5225139
f1_keywords:
- fm20.chm5225139
ms.prod: office
ms.assetid: 5b6c7176-0838-33da-1111-9591f961641e
ms.date: 11/14/2018
localization_priority: Normal
---


# InsideHeight, InsideWidth properties example

The following example uses the **[InsideHeight and InsideWidth](insideheight-insidewidth-properties.md)** properties to resize a **[CommandButton](commandbutton-control.md)**. The user clicks the **CommandButton** to resize it.

> [!NOTE] 
> **InsideHeight** and **InsideWidth** are read-only properties.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **CommandButton** named CommandButton1.
    

```vb
Dim Resize As Single 
 
Private Sub UserForm_Initialize() 
 Resize = 0.75 
 CommandButton1.Caption = "Resize Button" 
 
End Sub 
 
Private Sub CommandButton1_Click() 
 CommandButton1.Move 10, 10, _ 
 UserForm1.InsideWidth * Resize, _ 
 UserForm1.InsideHeight * Resize 
 CommandButton1.Caption = "Button resized " _ 
 & "using InsideHeight and InsideWidth!" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]