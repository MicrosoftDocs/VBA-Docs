---
title: Prevent the accidental erasure of data when moving between controls on a form
ms.prod: access
ms.assetid: 1733caa5-5067-e6d9-b614-51053180f22e
ms.date: 09/21/2018
localization_priority: Normal
---


# Prevent the accidental erasure of data when moving between controls on a form

When you tab from one text box or memo field to another in a form, the text in the control is highlighted. This makes it easy for users to accidentally delete the text by pressing a key. By using a few lines of code, you can move the insertion point to the first position in the text box, minimizing the risk of accidentally deleting the text. 

To do this, create a procedure for the text box's **[GotFocus](../../../api/Access.TextBox.GotFocus.md)** event. In the **GotFocus** event procedure, set the **[SelLength](../../../api/Access.TextBox.SelLength.md)** property of the text box to its **[SelStart](../../../api/Access.ComboBox.SelStart.md)** property. 

The following example illustrates how to do this for a text box named **txtFirstName**.

```vb
Private Sub txtFirstName_GotFocus() 
 
    Me.txtFirstName.SelLength = Me.txtFirstName.SelStart 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]