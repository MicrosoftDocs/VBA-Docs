---
title: "How to: Allow a Command Button to Take Focus Upon Clicking"
keywords: olfm10.chm3077253
f1_keywords:
- olfm10.chm3077253
ms.prod: outlook
ms.assetid: 7d4e4355-51cd-36cc-3e3c-18928f8cc03c
ms.date: 06/08/2017
localization_priority: Normal
---


# Allow a Command Button to Take Focus Upon Clicking

The following example uses the  **[TakeFocusOnClick](../../../api/Outlook.commandbutton.takefocusonclick.md)** property to control whether a **[CommandButton](../../../api/Outlook.commandbutton.md)** receives the focus when the user clicks on it. The user clicks a control other than CommandButton1 and then clicks CommandButton1. If **TakeFocusOnClick** is **True**, CommandButton1 receives the focus after it is clicked. The user can change the value of  **TakeFocusOnClick** by clicking the **[ToggleButton](../../../api/Outlook.togglebutton.md)**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **CommandButton** named CommandButton1.
    
- A  **ToggleButton** named ToggleButton1.
    
- One or two other controls, such as an  **[OptionButton](../../../api/Outlook.optionbutton.md)** or **[ListBox](../../../api/Outlook.listbox.md)**.
    



```vb
Sub CommandButton1_Click() 
 MsgBox "Watch CommandButton1 to see if it takes the focus." 
End Sub 
 
Sub ToggleButton1_Click() 
 Set ToggleButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ToggleButton1") 
 Set CommandButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("CommandButton1") 
 
 If ToggleButton1 = True Then 
 CommandButton1.TakeFocusOnClick = True 
 ToggleButton1.Caption = "TakeFocusOnClick On" 
 Else 
 CommandButton1.TakeFocusOnClick = False 
 ToggleButton1.Caption = "TakeFocusOnClick Off" 
 End If 
End Sub 
 
Sub Item_Open() 
 Set ToggleButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ToggleButton1") 
 Set CommandButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("CommandButton1") 
 
 CommandButton1.Caption = "Show Message" 
 
 ToggleButton1.Caption = "TakeFocusOnClick On" 
 ToggleButton1.Value = True 
 ToggleButton1.Width = 90 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]