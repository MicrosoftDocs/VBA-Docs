---
title: Declaring variables (VBA)
keywords: vbcn6.chm1076702
f1_keywords:
- vbcn6.chm1076702
ms.prod: office
ms.assetid: 42230f9e-e02f-14d9-8f7b-75441818e6c6
ms.date: 12/21/2018
localization_priority: Priority
---


# Declaring variables

When declaring [variables](../../Glossary/vbe-glossary.md#variable), you usually use a **[Dim](../../reference/user-interface-help/dim-statement.md)** statement. A declaration statement can be placed within a procedure to create a [procedure-level](../../Glossary/vbe-glossary.md#procedure-level) variable. Or it may be placed at the top of a [module](../../Glossary/vbe-glossary.md#module), in the Declarations section, to create a [module-level](../../Glossary/vbe-glossary.md#module-level) variable.

The following example creates the variable and specifies the [String data type](../../Glossary/vbe-glossary.md#string-data-type).

```vb
Dim strName As String 
```

If this statement appears within a procedure, the variable `strName` can be used only in that procedure. If the statement appears in the Declarations section of the module, the variable `strName` is available to all procedures within the module, but not to procedures in other modules in the [project](../../Glossary/vbe-glossary.md#project). 

To make this variable available to all procedures in the project, precede it with the **[Public](../../reference/user-interface-help/public-statement.md)** statement, as in the following example:

```vb
Public strName As String 
```

For information about naming your variables, see [Visual Basic naming rules](visual-basic-naming-rules.md).

Variables can be declared as one of the following [data types](../../reference/user-interface-help/data-type-summary.md): **Boolean**, **Byte**, **Integer**, **Long**, **Currency**, **Single**, **Double**, **Date**, **String** (for variable-length strings), **String * _length_** (for fixed-length strings), **Object**, or **Variant**. If you do not specify a data type, the **Variant** data type is assigned by default. You can also create a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type) by using the **[Type](../../reference/user-interface-help/type-statement.md)** statement. 

You can declare several variables in one statement. To specify a data type, you must include the data type for each variable. 

In the following statement, the variables `intX`, `intY`, and `intZ` are declared as type **Integer**.

```vb
Dim intX As Integer, intY As Integer, intZ As Integer 
```

In the following statement, `intX` and `intY` are declared as type **Variant**, and only `intZ` is declared as type **Integer**.

```vb
Dim intX, intY, intZ As Integer 
```

You don't have to supply the variable's data type in the declaration statement. If you omit the data type, the variable will be of type **Variant**.

## Public statement

You can use the **Public** statement to declare public module-level variables.

```vb
Public strName As String 
```

Public variables can be used in any procedures in the project. If a public variable is declared in a [standard module](../../Glossary/vbe-glossary.md#standard-module) or a [class module](../../Glossary/vbe-glossary.md#class-module), it can also be used in any projects that reference the project where the public variable is declared.


## Private statement

You can use the **[Private](../../reference/user-interface-help/private-statement.md)** statement to declare private module-level variables.

```vb
Private MyName As String 
```

Private variables can be used only by procedures in the same module.

> [!NOTE] 
> When used at the module level, the **Dim** statement is equivalent to the **Private** statement. You might want to use the **Private** statement to make your code easier to read and interpret.

## Static statement

When you use the **[Static](../../reference/user-interface-help/static-statement.md)** statement instead of a **Dim** statement to declare a variable in a procedure, the declared variable will retain its value between calls to that procedure.

## Option Explicit statement

You can implicitly declare a variable in Visual Basic simply by using it in an assignment statement. All variables that are implicitly declared are of type **Variant**. Variables of type **Variant** require more memory resources than most other variables. Your application will be more efficient if you declare variables explicitly and with a specific data type. Explicitly declaring all variables reduces the incidence of naming-conflict errors and spelling mistakes.

If you don't want Visual Basic to make implicit declarations, you can place the **[Option Explicit](../../reference/user-interface-help/option-explicit-statement.md)** statement in a module before any procedures. This statement requires you to explicitly declare all variables within the module. If a module includes the **Option Explicit** statement, a [compile-time](../../Glossary/vbe-glossary.md#compile-time) error will occur when Visual Basic encounters a variable name that has not been previously declared, or that has been spelled incorrectly.

You can set an option in your Visual Basic programming environment to automatically include the **Option Explicit** statement in all new modules. See your application's documentation for help on how to change Visual Basic environment options. Note that this option does not change existing code that you have written.

> [!NOTE] 
> You must explicitly declare fixed arrays and dynamic arrays.


## Declaring an object variable for automation

When you use one application to control another application's objects, you should set a reference to the other application's [type library](../../Glossary/vbe-glossary.md#type-library). After you set a reference, you can declare [object variables](../../Glossary/vbe-glossary.md#object-variable) according to their most specific type. For example, if you are in Microsoft Word when you set a reference to the Microsoft Excel type library, you can declare a variable of type **Worksheet** from within Word to represent an Excel **Worksheet** object.

If you are using another application to control Microsoft Access objects, in most cases, you can declare object variables according to their most specific type. You can also use the **New** keyword to create a new instance of an object automatically. However, you may have to indicate that it is a Microsoft Access object. For example, when you declare an object variable to represent an Access form from within Visual Basic, you must distinguish the Access **Form** object from a Visual Basic **Form** object. Include the name of the type library in the variable declaration, as in the following example:

```vb
Dim frmOrders As New Access.Form 
```

Some applications don't recognize individual Access object types. Even if you set a reference to the Access type library from these applications, you must declare all Access object variables as type **Object**. Nor can you use the **New** keyword to create a new instance of the object. 

The following example shows how to declare a variable to represent an instance of the Access **Application** object from an application that doesn't recognize Access object types. The application then creates an instance of the **Application** object.

```vb
Dim appAccess As Object 
Set appAccess = CreateObject("Access.Application")
```

To determine which syntax an application supports, see the application's documentation.


## See also

- [Data type summary](../../reference/user-interface-help/data-type-summary.md)
- [Variables and constants keyword summary](../../reference/user-interface-help/variables-and-constants-keyword-summary.md)
- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
