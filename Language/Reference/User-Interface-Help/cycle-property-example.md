---
title: Cycle property example
keywords: fm20.chm5225193
f1_keywords:
- fm20.chm5225193
ms.prod: office
ms.assetid: cf7a4e93-842e-5def-d7f7-214b6b37c180
ms.date: 11/14/2018
localization_priority: Normal
---


# Cycle property example

The following example defines the **[Cycle](cycle-property.md)** property for a **[Frame](frame-control.md)** and two **[Page](page-object.md)** objects in a **[MultiPage](multipage-control.md)**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **Frame** named Frame1.    
- A **MultiPage** named MultiPage1 that contains two objects named Page1 and Page2.    
- Two **[CommandButton](commandbutton-control.md)** controls named CommandButton1 and CommandButton2.
    
In the form, the **Frame** and each **Page** of the **MultiPage** place a couple of controls, so you can see how **Cycle** affects the tab order of the **Frame** and **MultiPage**.

The user should tab through the controls to observe how **Cycle** affects the tab order. Pressing CommandButton1 extends the tab order to include controls in the **Frame** and **Page** objects. Pressing CommandButton2 restricts the tab order.


```vb
Private Sub RestrictCycles() 
'Limit tab order for the Frame and Page objects 
 Frame1.Cycle = fmCycleCurrentForm 
 MultiPage1.Page1.Cycle = fmCycleCurrentForm 
 MultiPage1.Page2.Cycle = fmCycleCurrentForm 
End Sub 
 
Private Sub UserForm_Initialize() 
 RestrictCycles 
End Sub 
 
Private Sub CommandButton1_Click() 
'Extend tab order subforms (the Frame and Page 
'objects) 
 Frame1.Cycle = fmCycleAllForms 
 MultiPage1.Page1.Cycle = fmCycleAllForms 
 MultiPage1.Page2.Cycle = fmCycleAllForms 
End Sub 
 
Private Sub CommandButton2_Click() 
 RestrictCycles 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]