---
title: Name property example
keywords: fm20.chm5225156
f1_keywords:
- fm20.chm5225156
ms.prod: office
ms.assetid: d15fecd4-e195-3026-5c7c-5e0780f2f132
ms.date: 11/14/2018
localization_priority: Normal
---


# Name property example

The following example displays the **[Name](name-propertye-microsoft-forms.md)** property of each control on a form. This example uses the **[Controls](controls-collection-microsoft-forms.md)** collection to cycle through all the controls placed directly on the Userform.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains a **[CommandButton](commandbutton-control.md)** named CommandButton1 and several other controls.


```vb
Private Sub CommandButton1_Click() 
 Dim MyControl As Control 
 
 For Each MyControl In Controls 
 MsgBox "MyControl.Name = " & MyControl.Name 
 Next 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]