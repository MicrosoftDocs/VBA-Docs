---
title: Form.Form property (Access)
keywords: vbaac10.chm13499
f1_keywords:
- vbaac10.chm13499
ms.prod: access
api_name:
- Access.Form.Form
ms.assetid: 5e18dd48-f288-2b75-f42c-3a8b42f75b33
ms.date: 03/06/2019
localization_priority: Normal
---


# Form.Form property (Access)

You can use the **Form** property to refer to a form or to refer to the form associated with a subformcontrol. Read-only **Form**.


## Syntax

_expression_.**Form**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

This property refers to a form object. It is read-only in all views.

This property is typically used to refer to the form or report contained in a subform control. For example, the following code uses the **Form** property to access the **OrderID** control on a subform contained in the **OrderDetails** subform control.

```vb
Dim intOrderID As Integer 
intOrderID = Forms!Orders!OrderDetails.Form!OrderID
```

The next example calls a function from a property sheet by using the **Form** property to refer to the active form that contains the control named **CustomerID**.

```vb
=MyFunction(Form!CustomerID)
```

When you use the **Form** property in this manner, you are referring to the active form, and the name of the form isn't necessary.

The next example is the Visual Basic equivalent of the preceding example.

```vb
X = MyFunction(Forms!Customers!CustomerID)
```

> [!NOTE] 
> When you use the **[Forms](Access.Forms.md)** collection, you must specify the name of the form.


## Example

The following example uses the **Form** property to refer to a control on a subform.

```vb
Dim curTotalAmount As Currency 
 
curTotalAmount = Forms!Orders!OrderDetails.Form!TotalAmount 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]