---
title: Cancel the deletion of a record by using custom criteria
ms.prod: access
ms.assetid: 0445765f-4629-5970-776c-5bd30e2d72a1
ms.date: 09/25/2018
localization_priority: Normal
---


# Cancel the deletion of a record by using custom criteria

The following example illustrates how to use a form's **[Delete](../../../api/Access.Form.Delete.md)** event to prevent the deletion of a record based on custom criteria. In this example, the **Delete** event is canceled if the value of the DataRequired field is Yes.


```vb
Private Sub Form_Delete(Cancel As Integer) 
 
   ' Check the value of the DataRequired field. 
    If Me.DataRequired = "Yes" Then 
 
      ' Cancel the record deletion. 
      Cancel = True 
 
      ' Notify the user. 
       MsgBox "Cannot Delete the Record." 
    End If 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]