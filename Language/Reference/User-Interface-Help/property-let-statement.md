---
title: Property Let Statement
keywords: vblr6.chm1009540
f1_keywords:
- vblr6.chm1009540
ms.prod: office
ms.assetid: ecc8c277-ca44-add3-81c9-262219b1f7d6
ms.date: 06/08/2017
---


# Property Let Statement

Declares the name, [arguments](../../Glossary/vbe-glossary.md#argument), and code that form the body of a  **Property** **Let**[procedure](../../Glossary/vbe-glossary.md#procedure), which assigns a value to a [property](../../Glossary/vbe-glossary.md#property).

## Syntax

[ **Public** |**Private** |**Friend** ] [ **Static** ] **Property** **Let**_name_**(** [ _arglist_**,** ] _value_**)**
[ _statements_ ]
[ **Exit Property** ]
[ _statements_ ]

 **End Property**
The  **Property Let** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
|**Public**|Optional. Indicates that the  **Property** **Let** procedure is accessible to all other procedures in all[modules](../../Glossary/vbe-glossary.md#module). If used in a module that contains an  **Option Private** statement, the procedure is not available outside the[project](../../Glossary/vbe-glossary.md#project).|
|**Private**|Optional. Indicates that the  **Property** **Let** procedure is accessible only to other procedures in the module where it is declared.|
|**Friend**|Optional. Used only in a [class module](../../Glossary/vbe-glossary.md#class-module). Indicates that the  **Property Let** procedure is visible throughout the[project](../../Glossary/vbe-glossary.md#project), but not visible to a controller of an instance of an object.|
|**Static**|Optional. Indicates that the  **Property Let** procedure's local[variables](../../Glossary/vbe-glossary.md#variable) are preserved between calls. The **Static** attribute doesn't affect variables that are declared outside the **Property Let** procedure, even if they are used in the procedure.|
| _name_|Required. Name of the  **Property Let** procedure; follows standard variable naming conventions, except that the name can be the same as a **Property Get** or **Property Set** procedure in the same module.|
| _arglist_|Required. List of variables representing arguments that are passed to the  **Property Let** procedure when it is called. Multiple arguments are separated by commas. The name and[data type](../../Glossary/vbe-glossary.md#data-type) of each argument in a **Property Let** procedure must be the same as the corresponding argument in a **Property Get** procedure.|
| _value_|Required. Variable to contain the value to be assigned to the property. When the procedure is called, this argument appears on the right side of the calling [expression](../../Glossary/vbe-glossary.md#expression). The data type of  _value_ must be the same as the return type of the corresponding **Property Get** procedure.|
| _statements_|Optional. Any group of [statements](../../Glossary/vbe-glossary.md#statement) to be executed within the **Property Let** procedure.|

The  _arglist_ argument has the following syntax and parts:
[ **Optional** ] [ **ByVal** |**ByRef** ] [ **ParamArray** ] _varname_ [ **( )** ] [ **As**_type_ ] [ **=**_defaultvalue_ ]


|**Part**|**Description**|
|:-----|:-----|
|**Optional**|Optional. Indicates that an argument is not required. If used, all subsequent arguments in  _arglist_ must also be optional and declared using the **Optional** keyword. Note that it is not possible for the right side of a **Property Let** expression to be **Optional**.|
|**ByVal**|Optional. Indicates that the argument is passed [by value](../../Glossary/vbe-glossary.md#by-value).|
|**ByRef**|Optional. Indicates that the argument is passed [by reference](../../Glossary/vbe-glossary.md#by-reference).  **ByRef** is the default in Visual Basic.|
|**ParamArray**|Optional. Used only as the last argument in  _arglist_ to indicate that the final argument is an **Optional** array of **Variant** elements. The **ParamArray** keyword allows you to provide an arbitrary number of arguments. It may not be used with **ByVal**, **ByRef**, or **Optional**.|
| _varname_|Required. Name of the variable representing the argument; follows standard variable naming conventions.|
| _type_|Optional. Data type of the argument passed to the procedure; may be [Byte](../../Glossary/vbe-glossary.md#Byte), [Boolean](../../Glossary/vbe-glossary.md#Boolean), [Integer](../../Glossary/vbe-glossary.md#Integer), [Long](../../Glossary/vbe-glossary.md#Long), [Currency](../../Glossary/vbe-glossary.md#Currency), [Single](../../Glossary/vbe-glossary.md#Single), [Double](../../Glossary/vbe-glossary.md#Double), [Decimal](../../Glossary/vbe-glossary.md#Decimal) (not currently supported),[Date](../../Glossary/vbe-glossary.md#Date), [String](../../Glossary/vbe-glossary.md#String) (variable length only),[Object](../../Glossary/vbe-glossary.md#Object), [Variant](../../Glossary/vbe-glossary.md#Variant), or a specific [object type](../../Glossary/vbe-glossary.md#object-type). If the parameter is not  **Optional**, a[user-defined type](../../Glossary/vbe-glossary.md#user-defined-type) may also be specified.|
| _defaultvalue_|Optional. Any [constant](../../Glossary/vbe-glossary.md#constant) or constant expression. Valid for **Optional** parameters only. If the type is an **Object**, an explicit default value can only be **Nothing**.|

 **Note**  Every  **Property Let** statement must define at least one argument for the procedure it defines. That argument (or the last argument if there is more than one) contains the actual value to be assigned to the property when the procedure defined by the **Property Let** statement is invoked. That argument is referred to as _value_ in the preceding syntax.

## Remarks

If not explicitly specified using  **Public**, **Private**, or **Friend**, **Property** procedures are public by default. If **Static** isn't used, the value of local variables is not preserved between calls. The **Friend** keyword can only be used in class modules. However, **Friend** procedures can be accessed by procedures in any module of a project. A **Friend** procedure doesn't appear in the[type library](../../Glossary/vbe-glossary.md#type-library) of its parent class, nor can a **Friend** procedure be late bound.
All executable code must be in procedures. You can't define a  **Property Let** procedure inside another **Property**, **Sub**, or **Function** procedure.
The  **Exit Property** statement causes an immediate exit from a **Property Let** procedure. Program execution continues with the statement following the statement that called the **Property Let** procedure. Any number of **Exit Property** statements can appear anywhere in a **Property Let** procedure.
Like a  **Function** and **Property Get** procedure, a **Property Let** procedure is a separate procedure that can take arguments, perform a series of statements, and change the value of its arguments. However, unlike a **Function** and **Property Get** procedure, both of which return a value, you can only use a **Property Let** procedure on the left side of a property assignment expression or **Let** statement.

## Example

This example uses the  **Property Let** statement to define a procedure that assigns a value to a property. The property identifies the pen color for a drawing package.


```vb
Dim CurrentColor As Integer 
Const BLACK = 0, RED = 1, GREEN = 2, BLUE = 3 
 
' Set the pen color property for a Drawing package. 
' The module-level variable CurrentColor is set to 
' a numeric value that identifies the color used for drawing. 
Property Let PenColor(ColorName As String) 
 Select Case ColorName ' Check color name string. 
 Case "Red" 
 CurrentColor = RED ' Assign value for Red. 
 Case "Green" 
 CurrentColor = GREEN ' Assign value for Green. 
 Case "Blue" 
 CurrentColor = BLUE ' Assign value for Blue. 
 Case Else 
 CurrentColor = BLACK ' Assign default value. 
 End Select 
End Property 
 
' The following code sets the PenColor property for a drawing package 
' by calling the Property let procedure. 
 
PenColor = "Red" 

```


