---
title: Hide a subform if the main form contains no records
ms.prod: access
ms.assetid: 20482340-0c86-71c9-3ba1-b9f515397fbc
ms.date: 09/21/2018
localization_priority: Normal
---


# Hide a subform if the main form contains no records

The following example illustrates how to hide a subform named _Orders_Subform_ if its main form does not contain any records. The code resides in the main form's **[Current](../../../api/Access.Form.Current.md)** event procedure.


```vb
Private Sub Form_Current() 
 
    With Me![Orders_Subform].Form 
     
        ' Check the RecordCount of the Subform. 
        If .RecordsetClone.RecordCount = 0 Then 
         
            ' Hide the subform. 
            .Visible = False 
         
        End If 
    End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]