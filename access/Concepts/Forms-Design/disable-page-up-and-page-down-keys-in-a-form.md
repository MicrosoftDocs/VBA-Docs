---
title: Disable PAGE UP and PAGE DOWN keys in a form
ms.prod: access
ms.assetid: 998e1d00-f9d3-fcca-4535-390b0fd0d482
ms.date: 09/25/2018
localization_priority: Normal
---


# Disable PAGE UP and PAGE DOWN keys in a form

By default, the PAGE UP and PAGE DOWN keys can be used to navigate between records in a form. The following example illustrates how to use a form's **[KeyDown](../../../api/Access.Form.KeyDown.md)** event to disable the use of the PAGE UP and PAGE DOWN keys in the form.


```vb
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) 
 
    ' The Keycode value represents the key that 
    ' triggered the event. 
    Select Case KeyCode 
    
        ' Check for the PAGE UP and PAGE DOWN keys. 
        Case 33, 34 
 
        ' Cancel the keystroke. 
        KeyCode = 0 
    End Select 
End Sub
```


> [!NOTE] 
> You must set the form's **[KeyPreview](../../../api/Access.Form.KeyPreview.md)** property to **True** for this procedure to work.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]