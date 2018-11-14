---
title: Page Object, CommandButton, MultiPage Controls, ControlTipText Property Example
keywords: fm20.chm5225186
f1_keywords:
- fm20.chm5225186
ms.prod: office
ms.assetid: b7b8aac6-353c-1af9-de6b-e3de110c55ff
ms.date: 06/08/2017
---


# Page Object, CommandButton, MultiPage Controls, ControlTipText Property Example

The following example defines the  **ControlTipText** property for three **[CommandButton](commandbutton-control.md)** controls and two **Page** objects in a **[MultiPage](multipage-control.md)**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **[MultiPage](multipage-control.md)** named MultiPage1.
    
- Three  **[CommandButton](commandbutton-control.md)** controls named CommandButton1 through CommandButton3.
    


 **Note**  For an individual  **Page** of a **[MultiPage](multipage-control.md)**, **ControlTipText** becomes enabled when the **[MultiPage](multipage-control.md)** or a control on the current page of the **[MultiPage](multipage-control.md)** has the focus.




```vb
Private Sub UserForm_Initialize() 
 MultiPage1.Page1.ControlTipText = "Here in page 1" 
 MultiPage1.Page2.ControlTipText = "Now in page 2" 
 
 CommandButton1.ControlTipText = "And now here's" 
 CommandButton2.ControlTipText = "a tip from" 
 CommandButton3.ControlTipText = "your controls!" 
End Sub
```


