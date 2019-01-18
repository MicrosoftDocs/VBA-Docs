---
title: Page object, CommandButton, MultiPage controls, ControlTipText property example
keywords: fm20.chm5225186
f1_keywords:
- fm20.chm5225186
ms.prod: office
ms.assetid: b7b8aac6-353c-1af9-de6b-e3de110c55ff
ms.date: 11/14/2018
localization_priority: Normal
---


# Page object, CommandButton, MultiPage controls, ControlTipText property example

The following example defines the **[ControlTipText](controltiptext-property.md)** property for three **[CommandButton](commandbutton-control.md)** controls and two **[Page](page-object.md)** objects in a **[MultiPage](multipage-control.md)**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **MultiPage** named MultiPage1.   
- Three **CommandButton** controls named CommandButton1 through CommandButton3.
    
> [!NOTE] 
> For an individual **Page** of a **MultiPage**, **ControlTipText** becomes enabled when the **MultiPage** or a control on the current page of the **MultiPage** has the focus.


```vb
Private Sub UserForm_Initialize() 
 MultiPage1.Page1.ControlTipText = "Here in page 1" 
 MultiPage1.Page2.ControlTipText = "Now in page 2" 
 
 CommandButton1.ControlTipText = "And now here's" 
 CommandButton2.ControlTipText = "a tip from" 
 CommandButton3.ControlTipText = "your controls!" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]