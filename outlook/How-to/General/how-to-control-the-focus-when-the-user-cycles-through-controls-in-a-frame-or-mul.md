---
title: "How to: Control the Focus When the User Cycles through Controls in a Frame or MultiPage Control on a Form"
keywords: olfm10.chm3077172
f1_keywords:
- olfm10.chm3077172
ms.prod: outlook
ms.assetid: c7d1ac62-3c11-040a-d0f2-1f3e04c89f15
ms.date: 06/08/2017
localization_priority: Normal
---


# Control the Focus When the User Cycles through Controls in a Frame or MultiPage Control on a Form

The following example defines the **Cycle** property for a **[Frame](../../../api/Outlook.frame.md)** and two **[Page](../../../api/Outlook.page.md)** objects in a **[MultiPage](../../../api/Outlook.multipage.md)**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the **Open** event will activate. Make sure that the form contains:

- A **Frame** named Frame1.
    
- A **MultiPage** named MultiPage1 that contains two objects named Page1 and Page2.
    
- Two **[CommandButton](../../../api/Outlook.commandbutton.md)** controls named CommandButton1 and CommandButton2.
    
In the form, the **Frame**, and each **Page** of the **MultiPage**, place a couple of controls, so you can see how **Cycle** affects the tab order of the **Frame** and **MultiPage**.
The user should tab through the controls to observe how **Cycle** affects the tab order. Pressing CommandButton1 extends the tab order to include controls in the **Frame** and **Page** objects. Pressing CommandButton2 restricts the tab order.



```vb
Dim Frame1 
Dim MultiPage1 
 
Sub Item_Open() 
 Set Frame1 = Item.GetInspector.ModifiedFormPages("P.2").Frame1 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").MultiPage1 
 RestrictCycles 
End Sub 
 
Sub RestrictCycles() 
 'Limit tab order for the Frame and Page objects 
 Frame1.Cycle = 2 
 MultiPage1.Page1.Cycle = 2 
 MultiPage1.Page2.Cycle = 2 
End Sub 
Sub CommandButton1_Click() 
 'Extend tab order subforms (the Frame and Page objects) 
 Frame1.Cycle = 0 
 MultiPage1.Page1.Cycle = 0 
 MultiPage1.Page2.Cycle = 0 
End Sub 
 
Sub CommandButton2_Click() 
 RestrictCycles 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]