---
title: Function statement (VBA)
keywords: vblr6.chm1008927
f1_keywords:
- vblr6.chm1008927
ms.prod: office
ms.assetid: 407a6e70-b3e4-f13a-bda9-59296b288287
ms.date: 12/03/2018
localization_priority: Normal
---


# Function statement

Declares the name, [arguments](../../Glossary/vbe-glossary.md#argument), and code that form the body of a **Function** [procedure](../../Glossary/vbe-glossary.md#procedure).

## Syntax

[**Public** | **Private** | **Friend**] [ **Static** ] **Function** _name_ [ ( _arglist_ ) ] [ **As** _type_ ]<br/>
[ _statements_ ]<br/>
[ _name_ **=** _expression_ ]<br/>
[ **Exit Function** ]<br/>
[ _statements_ ]<br/>
[ _name_ **=** _expression_ ]<br/>
**End Function**

<br/>

The **Function** statement syntax has these parts:

|Part|Description|
|:-----|:-----|
|**Public**|Optional. Indicates that the **Function** procedure is accessible to all other procedures in all [modules](../../Glossary/vbe-glossary.md#module). If used in a module that contains an **Option Private**, the procedure is not available outside the [project](../../Glossary/vbe-glossary.md#project).|
|**Private**|Optional. Indicates that the **Function** procedure is accessible only to other procedures in the module where it is declared.|
|**Friend**|Optional. Used only in a [class module](../../Glossary/vbe-glossary.md#class-module). Indicates that the **Function** procedure is visible throughout the project, but not visible to a controller of an instance of an object.|
|**Static**|Optional. Indicates that the **Function** procedure's local [variables](../../Glossary/vbe-glossary.md#variable) are preserved between calls. The **Static** attribute doesn't affect variables that are declared outside the **Function**, even if they are used in the procedure.|
| _name_|Required. Name of the **Function**; follows standard variable naming conventions.|
| _arglist_|Optional. List of variables representing arguments that are passed to the **Function** procedure when it is called. Multiple variables are separated by commas.|
| _type_|Optional. [Data type](../../Glossary/vbe-glossary.md#data-type) of the value returned by the **Function** procedure; may be [Byte](../../Glossary/vbe-glossary.md#byte-data-type), [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type), [Integer](../../Glossary/vbe-glossary.md#integer-data-type), [Long](../../Glossary/vbe-glossary.md#long-data-type), [Currency](../../Glossary/vbe-glossary.md#currency-data-type), [Single](../../Glossary/vbe-glossary.md#single-data-type), [Double](../../Glossary/vbe-glossary.md#double-data-type), [Decimal](../../Glossary/vbe-glossary.md#decimal-data-type) (not currently supported), [Date](../../Glossary/vbe-glossary.md#date-data-type), [String](../../Glossary/vbe-glossary.md#string-data-type) (except fixed length), [Object](../../Glossary/vbe-glossary.md#object), [Variant](../../Glossary/vbe-glossary.md#variant-data-type), or any [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type).|
| _statements_|Optional. Any group of statements to be executed within the **Function** procedure.|
| _expression_|Optional. Return value of the **Function**.|

<br/>

The _arglist_ argument has the following syntax and parts:

[ **Optional** ] [ **ByVal** | **ByRef** ] [ **ParamArray** ] _varname_ [ ( ) ] [ **As** _type_ ] [ **=** _defaultvalue_ ]

<br/>

|Part|Description|
|:-----|:-----|
|**Optional**|Optional. Indicates that an argument is not required. If used, all subsequent arguments in _arglist_ must also be optional and declared by using the **Optional** keyword. **Optional** can't be used for any argument if **ParamArray** is used.|
|**ByVal**|Optional. Indicates that the argument is passed [by value](../../Glossary/vbe-glossary.md#by-value).|
|**ByRef**|Optional. Indicates that the argument is passed [by reference](../../Glossary/vbe-glossary.md#by-reference). **ByRef** is the default in Visual Basic.|
|**ParamArray**|Optional. Used only as the last argument in _arglist_ to indicate that the final argument is an **Optional** array of **Variant** elements. The **ParamArray** keyword allows you to provide an arbitrary number of arguments. It may not be used with **ByVal**, **ByRef**, or **Optional**.|
| _varname_|Required. Name of the variable representing the argument; follows standard variable naming conventions.|
| _type_|Optional. Data type of the argument passed to the procedure; may be **Byte**, **Boolean**, **Integer**, **Long**, **Currency**, **Single**, **Double**, **Decimal** (not currently supported) **Date**, **String** (variable length only), **Object**, **Variant**, or a specific [object type](../../Glossary/vbe-glossary.md#object-type). If the parameter is not **Optional**, a user-defined type may also be specified.|
| _defaultvalue_|Optional. Any [constant](../../Glossary/vbe-glossary.md#constant) or constant expression. Valid for **Optional** parameters only. If the type is an **Object**, an explicit default value can only be **Nothing**.|

## Remarks

If not explicitly specified by using **Public**, **Private**, or **[Friend](friend-keyword.md)**, **Function** procedures are public by default. 

If **[Static](static-statement.md)** isn't used, the value of local variables is not preserved between calls. 

The **Friend** keyword can only be used in class modules. However, **Friend** procedures can be accessed by procedures in any module of a project. A **Friend** procedure does not appear in the [type library](../../Glossary/vbe-glossary.md#type-library) of its parent class, nor can a **Friend** procedure be late bound.

**Function** procedures can be recursive; that is, they can call themselves to perform a given task. However, recursion can lead to stack overflow. The **Static** keyword usually isn't used with recursive **Function** procedures.

All executable code must be in procedures. You can't define a **Function** procedure inside another **Function**, **[Sub](sub-statement.md)**, or **Property** procedure.

The **[Exit Function](exit-statement.md)** statement causes an immediate exit from a **Function** procedure. Program execution continues with the statement following the statement that called the **Function** procedure. Any number of **Exit Function** statements can appear anywhere in a **Function** procedure.

Like a **Sub** procedure, a **Function** procedure is a separate procedure that can take arguments, perform a series of statements, and change the values of its arguments. However, unlike a **Sub** procedure, you can use a **Function** procedure on the right side of an [expression](../../Glossary/vbe-glossary.md#expression) in the same way you use any intrinsic function, such as **Sqr**, **Cos**, or **Chr**, when you want to use the value returned by the function.

You call a **Function** procedure by using the function name, followed by the argument list in parentheses, in an expression. See the **[Call](call-statement.md)** statement for specific information about how to call **Function** procedures.

To return a value from a function, assign the value to the function name. Any number of such assignments can appear anywhere within the procedure. If no value is assigned to _name_, the procedure returns a default value: a numeric function returns 0, a string function returns a zero-length string (""), and a **Variant** function returns [Empty](../../Glossary/vbe-glossary.md#empty). A function that returns an object reference returns **Nothing** if no object reference is assigned to _name_ (using **Set**) within the **Function**.

The following example shows how to assign a return value to a function. In this case, **False** is assigned to the name to indicate that some value was not found.

```vb
Function BinarySearch(. . .) As Boolean 
'. . . 
 ' Value not found. Return a value of False. 
 If lower > upper Then 
  BinarySearch = False 
  Exit Function 
 End If 
'. . . 
End Function
```

Variables used in **Function** procedures fall into two categories: those that are explicitly declared within the procedure and those that are not. 

Variables that are explicitly declared in a procedure (using **Dim** or the equivalent) are always local to the procedure. Variables that are used but not explicitly declared in a procedure are also local unless they are explicitly declared at some higher level outside the procedure.

A procedure can use a variable that is not explicitly declared in the procedure, but a naming conflict can occur if anything you defined at the [module level](../../Glossary/vbe-glossary.md#module-level) has the same name. If your procedure refers to an undeclared variable that has the same name as another procedure, constant, or variable, it is assumed that your procedure refers to that module-level name. Explicitly declare variables to avoid this kind of conflict. You can use an **[Option Explicit](option-explicit-statement.md)** statement to force explicit declaration of variables.

Visual Basic may rearrange arithmetic expressions to increase internal efficiency. Avoid using a **Function** procedure in an arithmetic expression when the function changes the value of variables in the same expression. For more information about arithmetic operators, see [Operators](operator-summary.md).

## Example

This example uses the **Function** statement to declare the name, arguments, and code that form the body of a **Function** procedure. The last example uses hard-typed, initialized **Optional** arguments.

```vb
' The following user-defined function returns the square root of the 
' argument passed to it. 
Function CalculateSquareRoot(NumberArg As Double) As Double 
 If NumberArg < 0 Then ' Evaluate argument. 
  Exit Function ' Exit to calling procedure. 
 Else 
  CalculateSquareRoot = Sqr(NumberArg) ' Return square root. 
 End If 
End Function
```

<br/>

Using the **ParamArray** keyword enables a function to accept a variable number of arguments. In the following definition, it is passed by value.

```vb
Function CalcSum(ByVal FirstArg As Integer, ParamArray OtherArgs()) 
Dim ReturnValue 
' If the function is invoked as follows: 
ReturnValue = CalcSum(4, 3, 2, 1) 
' Local variables are assigned the following values: FirstArg = 4, 
' OtherArgs(1) = 3, OtherArgs(2) = 2, and so on, assuming default 
' lower bound for arrays = 1. 

```

<br/>

**Optional** arguments can have default values and types other than **Variant**.

```vb
' If a function's arguments are defined as follows: 
Function MyFunc(MyStr As String,Optional MyArg1 As _
 Integer = 5,Optional MyArg2 = "Dolly") 
Dim RetVal 
' The function can be invoked as follows: 
RetVal = MyFunc("Hello", 2, "World") ' All 3 arguments supplied. 
RetVal = MyFunc("Test", , 5) ' Second argument omitted. 
' Arguments one and three using named-arguments. 
RetVal = MyFunc(MyStr:="Hello ", MyArg1:=7) 

```


## See also

- [Calling Sub and Function procedures](../../concepts/getting-started/calling-sub-and-function-procedures.md)
- [Understanding named arguments and optional arguments](../../concepts/getting-started/understanding-named-arguments-and-optional-arguments.md)
- [Writing a Function procedure](../../concepts/getting-started/writing-a-function-procedure.md)
- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
