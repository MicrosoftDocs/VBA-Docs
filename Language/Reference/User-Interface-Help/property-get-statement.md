---
title: Property Get statement (VBA)
keywords: vblr6.chm1009538
f1_keywords:
- vblr6.chm1009538
ms.prod: office
ms.assetid: 39d1fb20-653e-a174-7a98-e2b33f260d39
ms.date: 12/03/2018
localization_priority: Normal
---


# Property Get statement

Declares the name, [arguments](../../Glossary/vbe-glossary.md#argument), and code that form the body of a **Property** [procedure](../../Glossary/vbe-glossary.md#procedure), which gets the value of a [property](../../Glossary/vbe-glossary.md#property).

## Syntax

[ **Public** | **Private** | **Friend** ] [ **Static** ] **Property Get**_name_ [ (_arglist_) ] [ **As** _type_ ] <br/>
[ _statements_ ] <br/>
[ _name_ **=** _expression_ ] <br/>
[ **Exit Property** ] <br/>
[ _statements_ ] <br/>
[ _name_ **=** _expression_ ] <br/>
**End Property**

<br/>

The **Property Get** statement syntax has these parts:

|Part|Description|
|:-----|:-----|
|**Public**|Optional. Indicates that the **Property Get** procedure is accessible to all other procedures in all [modules](../../Glossary/vbe-glossary.md#module). If used in a module that contains an **Option Private** statement, the procedure is not available outside the [project](../../Glossary/vbe-glossary.md#project).|
|**Private**|Optional. Indicates that the **Property Get** procedure is accessible only to other procedures in the module where it is declared.|
|**Friend**|Optional. Used only in a [class module](../../Glossary/vbe-glossary.md#class-module). Indicates that the **Property Get** procedure is visible throughout the project, but not visible to a controller of an instance of an object.|
|**Static**|Optional. Indicates that the **Property Get** procedure's local [variables](../../Glossary/vbe-glossary.md#variable) are preserved between calls. The **Static** attribute doesn't affect variables that are declared outside the **Property Get** procedure, even if they are used in the procedure.|
| _name_|Required. Name of the **Property Get** procedure; follows standard variable naming conventions, except that the name can be the same as a **[Property Let](property-let-statement.md)** or **[Property Set](property-set-statement.md)** procedure in the same module.|
| _arglist_|Optional. List of variables representing arguments that are passed to the **Property Get** procedure when it is called. Multiple arguments are separated by commas. The name and [data type](../../Glossary/vbe-glossary.md#data-type) of each argument in a **Property Get** procedure must be the same as the corresponding argument in a **Property Let** procedure (if one exists).|
| _type_|Optional. Data type of the value returned by the **Property Get** procedure; may be [Byte](../../Glossary/vbe-glossary.md#byte-data-type), [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type), [Integer](../../Glossary/vbe-glossary.md#integer-data-type), [Long](../../Glossary/vbe-glossary.md#long-data-type), [Currency](../../Glossary/vbe-glossary.md#currency-data-type), [Single](../../Glossary/vbe-glossary.md#single-data-type), [Double](../../Glossary/vbe-glossary.md#double-data-type), [Decimal](../../Glossary/vbe-glossary.md#decimal-data-type) (not currently supported), [Date](../../Glossary/vbe-glossary.md#date-data-type), [String](../../Glossary/vbe-glossary.md#string-data-type) (except fixed length), [Object](../../Glossary/vbe-glossary.md#object), [Variant](../../Glossary/vbe-glossary.md#variant-data-type), [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type), and [Arrays](../../Glossary/vbe-glossary.md#array). <br/><br/>The return _type_ of a **Property Get** procedure must be the same data type as the last (or sometimes the only) argument in a corresponding **Property Let** procedure (if one exists) that defines the value assigned to the property on the right side of an [expression](../../Glossary/vbe-glossary.md#expression).|
| _statements_|Optional. Any group of statements to be executed within the body of the **Property Get** procedure.|
| _expression_|Optional. Value of the property returned by the procedure defined by the **Property Get** statement.|

<br/>

The _arglist_ argument has the following syntax and parts:

[ **Optional** ] [ **ByVal** | **ByRef** ] [ **ParamArray** ] _varname_ [ ( ) ] [ **As** _type_ ] [ **=** _defaultvalue_ ]

<br/>

|Part|Description|
|:-----|:-----|
|**Optional**|Optional. Indicates that an argument is not required. If used, all subsequent arguments in _arglist_ must also be optional and declared by using the **Optional** keyword.|
|**ByVal**|Optional. Indicates that the argument is passed [by value](../../Glossary/vbe-glossary.md#by-value).|
|**ByRef**|Optional. Indicates that the argument is passed [by reference](../../Glossary/vbe-glossary.md#by-reference). **ByRef** is the default in Visual Basic.|
|**ParamArray**|Optional. Used only as the last argument in _arglist_ to indicate that the final argument is an **Optional** array of **Variant** elements. The **ParamArray** keyword allows you to provide an arbitrary number of arguments. It may not be used with **ByVal**, **ByRef**, or **Optional**.|
| _varname_|Required. Name of the variable representing the argument; follows standard variable naming conventions.|
| _type_|Optional. Data type of the argument passed to the procedure; may be **Byte**, **Boolean**, **Integer**, **Long**, **Currency**, **Single**, **Double**, **Decimal** (not currently supported), **Date**, **String** (variable length only), **Object**, **Variant**, or a specific [object type](../../Glossary/vbe-glossary.md#object-type). If the parameter is not **Optional**, a user-defined type may also be specified.|
| _defaultvalue_|Optional. Any [constant](../../Glossary/vbe-glossary.md#constant) or constant expression. Valid for **Optional** parameters only. If the type is an **Object**, an explicit default value can only be **[Nothing](nothing-keyword.md)**.|

## Remarks

If not explicitly specified by using **[Public](public-statement.md)**, **[Private](private-statement.md)**, or **[Friend](friend-keyword.md)**, **Property** procedures are public by default. If **[Static](static-statement.md)** is not used, the value of local variables is not preserved between calls. 

The **Friend** keyword can only be used in class modules. However, **Friend** procedures can be accessed by procedures in any module of a project. A **Friend** procedure doesn't appear in the [type library](../../Glossary/vbe-glossary.md#type-library) of its parent class, nor can a **Friend** procedure be late bound.

All executable code must be in procedures. You can't define a **Property Get** procedure inside another **Property**, **[Sub](sub-statement.md)**, or **[Function](function-statement.md)** procedure.

The **[Exit Property](exit-statement.md)** statement causes an immediate exit from a **Property Get** procedure. Program execution continues with the statement following the statement that called the **Property Get** procedure. Any number of **Exit Property** statements can appear anywhere in a **Property Get** procedure.

Like a **Sub** and **Property Let** procedure, a **Property Get** procedure is a separate procedure that can take arguments, perform a series of statements, and change the values of its arguments. However, unlike a **Sub** or **Property Let** procedure, you can use a **Property Get** procedure on the right side of an expression in the same way that you use a **Function** or a property name when you want to return the value of a property.

## Example

This example uses the **Property Get** statement to define a property procedure that gets the value of a property. The property identifies the current color of a pen as a string.


```vb
Dim CurrentColor As Integer 
Const BLACK = 0, RED = 1, GREEN = 2, BLUE = 3 
 
' Returns the current color of the pen as a string. 
Property Get PenColor() As String 
 Select Case CurrentColor 
 Case RED 
 PenColor = "Red" 
 Case GREEN 
 PenColor = "Green" 
 Case BLUE 
 PenColor = "Blue" 
 End Select 
End Property 
 
' The following code gets the color of the pen 
' calling the Property Get procedure. 
ColorName = PenColor 

```

## See also

- [Calling property procedures](../../concepts/getting-started/calling-property-procedures.md)
- [Executing code when setting properties](../../concepts/getting-started/executing-code-when-setting-properties.md)
- [Writing a property procedure](../../concepts/getting-started/writing-a-property-procedure.md)
- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
