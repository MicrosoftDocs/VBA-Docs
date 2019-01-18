---
title: Executing code when setting properties (VBA)
keywords: vbcn6.chm1101378
f1_keywords:
- vbcn6.chm1101378
ms.prod: office
ms.assetid: ff32c6d2-1857-102f-371e-2d0f6ab848dc
ms.date: 12/21/2018
localization_priority: Normal
---


# Executing code when setting properties

You can create **[Property Let](../../reference/user-interface-help/property-let-statement.md)**, **[Property Set](../../reference/user-interface-help/property-set-statement.md)**, and **[Property Get](../../reference/user-interface-help/property-get-statement.md)** procedures that share the same name. By doing this, you can create a group of related [procedures](../../Glossary/vbe-glossary.md#procedure) that work together. After a name is used for a **Property** procedure, that name can't be used to name a **[Sub](../../reference/user-interface-help/sub-statement.md)** or **[Function](../../reference/user-interface-help/function-statement.md)** procedure, a [variable](../../Glossary/vbe-glossary.md#variable), or a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type).

The **Property Let** statement allows you to create a procedure that sets the value of the [property](../../Glossary/vbe-glossary.md#property). One example might be a **Property** procedure that creates an inverted property for a bitmap on a form. 

This is the syntax used to call the **Property Let** procedure.

```vb
Form1.Inverted = True 

```

The actual work of inverting a bitmap on the form is done within the **Property Let** procedure.

```vb
Private IsInverted As Boolean 
 
Property Let Inverted(X As Boolean) 
 IsInverted = X 
 If IsInverted Then 
 … 
 (statements) 
 Else 
 (statements) 
 End If 
End Property 

```

The form-level variable stores the setting of your property. By declaring it **Private**, the user can only change it by using your **Property Let** procedure. Use a name that makes it easy to recognize that the variable is used for the property.

This **Property Get** procedure is used to return the current state of the property.

```vb
Property Get Inverted() As Boolean 
 Inverted = IsInverted 
End Property 

```

[Property procedures](../../Glossary/vbe-glossary.md#property-procedure) make it easy to execute code at the same time that the value of a property is set. You can use property procedures to do the following processing:

- Before a property value is set to determine the value of the property.   
- After a property value is set, based on the new value.
    
## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]