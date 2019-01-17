---
title: Call procedures in a subform or subreport
ms.prod: access
ms.assetid: d0128a6c-f85b-fbf0-22cb-bfd4a8eca3c8
ms.date: 09/21/2018
localization_priority: Normal
---


# Call procedures in a subform or subreport

You can call a procedure in a module associated with a subform or subreport in one of two ways. If the form containing the subform is open in Form view, you can refer to the procedure as a method on the subform. 

The following example shows how to call the procedure GetProductID in the Orders subform, which is bound to a subform control on the Orders form:

In the Orders Subform class module, enter this code:

```vb
Public Function GetProductID() As Variant 
 ' Return current productID. 
 GetProductID = ProductID 
End Function 
```

The following code illustrates how to call the GetProductID procedure.

```vb
Forms!Orders![Orders Subform].Form.GetProductID
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]