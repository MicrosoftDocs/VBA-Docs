---
title: Sub statement (VBA)
keywords: vblr6.chm1009038
f1_keywords:
- vblr6.chm1009038
ms.prod: office
ms.assetid: 7931d739-a61a-78ba-5b33-960c1bf908ce
ms.date: 12/03/2018
localization_priority: Normal
---


# Sub statement

Declares the name, [arguments](../../Glossary/vbe-glossary.md#argument), and code that form the body of a **Sub** [procedure](../../Glossary/vbe-glossary.md#procedure).

## Syntax

[ **Private** | **Public** | **Friend** ] [ **Static** ] **Sub** _name_ [ ( _arglist_ ) ] <br/>
[ _statements_ ] <br/>
[ **Exit Sub** ] <br/>
[ _statements_ ] <br/>
**End Sub**

<br/>

The **Sub** statement syntax has these parts:

|Part|Description|
|:-----|:-----|
|**Public**|Optional. Indicates that the **Sub** procedure is accessible to all other procedures in all [modules](../../Glossary/vbe-glossary.md#module). If used in a module that contains an **Option Private** statement, the procedure is not available outside the [project](../../Glossary/vbe-glossary.md#project).|
|**Private**|Optional. Indicates that the **Sub** procedure is accessible only to other procedures in the module where it is declared.|
|**Friend**|Optional. Used only in a [class module](../../Glossary/vbe-glossary.md#class-module). Indicates that the **Sub** procedure is visible throughout the [project](../../Glossary/vbe-glossary.md#project), but not visible to a controller of an instance of an object.|
|**Static**|Optional. Indicates that the **Sub** procedure's local [variables](../../Glossary/vbe-glossary.md#variable) are preserved between calls. The **Static** attribute doesn't affect variables that are declared outside the **Sub**, even if they are used in the procedure.|
| _name_|Required. Name of the **Sub**; follows standard [variable](../../Glossary/vbe-glossary.md#variable) naming conventions.|
| _arglist_|Optional. List of variables representing arguments that are passed to the **Sub** procedure when it is called. Multiple variables are separated by commas.|
| _statements_|Optional. Any group of [statements](../../Glossary/vbe-glossary.md#statement) to be executed within the **Sub** procedure.|

<br/>

The _arglist_ argument has the following syntax and parts:

[ **Optional** ] [ **ByVal** | **ByRef** ] [ **ParamArray** ] _varname_ [ ( ) ] [ **As** _type_ ] [ **=** _defaultvalue_ ]

<br/>

|Part|Description|
|:-----|:-----|
|**Optional**|Optional. [Keyword](../../Glossary/vbe-glossary.md#keyword) indicating that an argument is not required. If used, all subsequent arguments in _arglist_ must also be optional and declared by using the **Optional** keyword. **Optional** can't be used for any argument if **ParamArray** is used.|
|**ByVal**|Optional. Indicates that the argument is passed [by value](../../Glossary/vbe-glossary.md#by-value).|
|**ByRef**|Optional. Indicates that the argument is passed [by reference](../../Glossary/vbe-glossary.md#by-reference). **ByRef** is the default in Visual Basic.|
|**ParamArray**|Optional. Used only as the last argument in _arglist_ to indicate that the final argument is an **Optional** [array](../../Glossary/vbe-glossary.md#array) of **Variant** elements. The **ParamArray** keyword allows you to provide an arbitrary number of arguments. **ParamArray** can't be used with **ByVal**, **ByRef**, or **Optional**.|
| _varname_|Required. Name of the variable representing the argument; follows standard variable naming conventions.|
| _type_|Optional. [Data type](../../Glossary/vbe-glossary.md#data-type) of the argument passed to the procedure; may be [Byte](../../Glossary/vbe-glossary.md#byte-data-type), [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type), [Integer](../../Glossary/vbe-glossary.md#integer-data-type), [Long](../../Glossary/vbe-glossary.md#long-data-type), [Currency](../../Glossary/vbe-glossary.md#currency-data-type), [Single](../../Glossary/vbe-glossary.md#single-data-type), [Double](../../Glossary/vbe-glossary.md#double-data-type), [Decimal](../../Glossary/vbe-glossary.md#decimal-data-type) (not currently supported), [Date](../../Glossary/vbe-glossary.md#date-data-type), [String](../../Glossary/vbe-glossary.md#string-data-type) (variable-length only), [Object](../../Glossary/vbe-glossary.md#object), [Variant](../../Glossary/vbe-glossary.md#variant-data-type), or a specific [object type](../../Glossary/vbe-glossary.md#object-type). If the parameter is not **Optional**, a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type) may also be specified.|
| _defaultvalue_|Optional. Any [constant](../../Glossary/vbe-glossary.md#constant) or constant [expression](../../Glossary/vbe-glossary.md#expression). Valid for **Optional** parameters only. If the type is an **Object**, an explicit default value can only be **Nothing**.|

## Remarks

If not explicitly specified by using **Public**, **Private**, or **[Friend](friend-keyword.md)**, **Sub** procedures are public by default.

If **[Static](static-statement.md)** isn't used, the value of local variables is not preserved between calls.

The **Friend** keyword can only be used in class modules. However, **Friend** procedures can be accessed by procedures in any module of a project. A **Friend** procedure doesn't appear in the [type library](../../Glossary/vbe-glossary.md#type-library) of its parent class, nor can a **Friend** procedure be late bound.

**Sub** procedures can be recursive; that is, they can call themselves to perform a given task. However, recursion can lead to stack overflow. The **Static** keyword usually is not used with recursive **Sub** procedures.

All executable code must be in [procedures](../../Glossary/vbe-glossary.md#procedure). You can't define a **Sub** procedure inside another **Sub**, **Function**, or **Property** procedure.

The **Exit Sub** keywords cause an immediate exit from a **Sub** procedure. Program execution continues with the statement following the statement that called the **Sub** procedure. Any number of **Exit Sub** statements can appear anywhere in a **Sub** procedure.

Like a **Function** procedure, a **Sub** procedure is a separate procedure that can take arguments, perform a series of statements, and change the value of its arguments. However, unlike a **Function** procedure, which returns a value, a **Sub** procedure can't be used in an expression.

You call a **Sub** procedure by using the procedure name followed by the argument list. See the **[Call](call-statement.md)** statement for specific information about how to call **Sub** procedures.

Variables used in **Sub** procedures fall into two categories: those that are explicitly declared within the procedure and those that are not. Variables that are explicitly declared in a procedure (using **Dim** or the equivalent) are always local to the procedure. Variables that are used but not explicitly declared in a procedure are also local unless they are explicitly declared at some higher level outside the procedure.

A procedure can use a variable that is not explicitly declared in the procedure, but a naming conflict can occur if anything you defined at the [module level](../../Glossary/vbe-glossary.md#module-level) has the same name. If your procedure refers to an undeclared variable that has the same name as another procedure, constant or variable, it is assumed that your procedure is referring to that module-level name. To avoid this kind of conflict, explicitly declare variables. You can use an **Option Explicit** statement to force explicit declaration of variables.

> [!NOTE] 
> You can't use **GoSub**, **GoTo**, or **Return** to enter or exit a **Sub** procedure.

## Example

This example uses the **Sub** statement to define the name, arguments, and code that form the body of a **Sub** procedure.

```vb
' Sub procedure definition. 
' Sub procedure with two arguments. 
Sub SubComputeArea(Length, TheWidth) 

   Dim Area As Double ' Declare local variable. 

   If Length = 0 Or TheWidth = 0 Then 
      ' If either argument = 0. 
      Exit Sub ' Exit Sub immediately. 
   End If 
   
   Area = Length * TheWidth ' Calculate area of rectangle. 
   Debug.Print Area ' Print Area to Debug window. 

End Sub
```

## See also

- [Calling Sub and Function procedures](../../concepts/getting-started/calling-sub-and-function-procedures.md)
- [Understanding named arguments and optional arguments](../../concepts/getting-started/understanding-named-arguments-and-optional-arguments.md)
- [Writing a Sub procedure](../../concepts/getting-started/writing-a-sub-procedure.md)
- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
