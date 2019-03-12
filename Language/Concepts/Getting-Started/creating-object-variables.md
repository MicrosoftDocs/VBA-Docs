---
title: Creating object variables (VBA)
keywords: vbcn6.chm1011337
f1_keywords:
- vbcn6.chm1011337
ms.prod: office
ms.assetid: 6cff962e-4a3e-dfc3-8491-d31a308b1c55
ms.date: 12/21/2018
localization_priority: Normal
---


# Creating object variables

You can treat an [object variable](../../Glossary/vbe-glossary.md#object-variable) exactly the same as the [object](../../Glossary/vbe-glossary.md#object) to which it refers. You can set or return the [properties](../../Glossary/vbe-glossary.md#property) of the object or use any of its [methods](../../Glossary/vbe-glossary.md#method).

## Create an object variable

1. Declare the object variable.
    
2. Assign the object variable to an object.
    

## Declare an object variable

Use the **[Dim](../../reference/user-interface-help/dim-statement.md)** statement or one of the other declaration statements (**Public**, **Private**, or **Static**) to declare an object variable. A [variable](../../Glossary/vbe-glossary.md#variable) that refers to an object must be a **[Variant](../../reference/user-interface-help/variant-data-type.md)**, an **[Object](../../reference/user-interface-help/object-data-type.md)**, or a specific type of object. For example, the following declarations are valid:

```vb
' Declare MyObject as Variant data type. 
Dim MyObject 
' Declare MyObject as Object data type. 
Dim MyObject As Object 
' Declare MyObject as Font type. 
Dim MyObject As Font 

```

> [!NOTE] 
> If you use an object variable without declaring it first, the [data type](../../Glossary/vbe-glossary.md#data-type) of the object variable is **Variant** by default.

You can declare an object variable with the **Object** data type when the specific [object type](../../Glossary/vbe-glossary.md#object-type) is not known until the procedure runs. Use the **Object** data type to create a generic reference to any object.

If you know the specific object type, you should declare the object variable as that object type. For example, if the application contains a Sample object type, you can declare an object variable for that object by using either of these statements:

```vb
Dim MyObject As Object ' Declared as generic object. 
Dim MyObject As Sample ' Declared only as Sample object. 

```

Declaring specific object types provides automatic type checking, faster code, and improved readability.

## Assign an object variable to an object

Use the **[Set](../../reference/user-interface-help/set-statement.md)** statement to assign an object to an object variable. You can assign an [object expression](../../Glossary/vbe-glossary.md#object-expression) or **[Nothing](../../reference/user-interface-help/nothing-keyword.md)**. For example, the following object variable assignments are valid.

```vb
Set MyObject = YourObject ' Assign object reference. 
Set MyObject = Nothing ' Discontinue association. 

```

You can combine declaring an object variable with assigning an object to it by using the **New** [keyword](../../Glossary/vbe-glossary.md#keyword) with the **Set** statement. For example:

```vb
Set MyObject = New Object ' Create and Assign 

```

Setting an object variable equal to **Nothing** discontinues the association of the object variable with any specific object. This prevents you from accidentally changing the object by changing the variable. An object variable is always set to **Nothing** after closing the associated object so you can test whether the object variable points to a valid object. For example:

```vb
If Not MyObject Is Nothing Then 
 ' Variable refers to valid object. 
 . . . 
End If 

```

Of course, this test can never determine with absolute certainty whether a user has closed the application containing the object to which the object variable refers.

## Refer to the current instance of an object

Use the **[Me](../../reference/user-interface-help/me-keyword.md)** keyword to refer to the current instance of the object where the code is running. All procedures associated with the current object have access to the object referred to as **Me**. Using **Me** is particularly useful for passing information about the current instance of an object to a procedure in another module. For example, suppose you have the following procedure in a module:

```vb
Sub ChangeObjectColor(MyObjectName As Object) 
 MyObjectName.BackColor = RGB(Rnd * 256, Rnd * 256, Rnd * 256) 
End Sub
```

You can call the procedure and pass the current instance of the object as an argument by using the following statement:

```vb
ChangeObjectColor Me 
```

<br/>

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
