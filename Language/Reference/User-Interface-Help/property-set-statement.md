---
title: Property Set Statement
keywords: vblr6.chm1009539
f1_keywords:
- vblr6.chm1009539
ms.prod: office
ms.assetid: 462c3a14-bd67-eed7-9b5b-396283952b0b
ms.date: 06/08/2017
---


# Property Set Statement

Declares the name, [arguments](../../Glossary/vbe-glossary.md#argument), and code that form the body of a  **Property**[procedure](../../Glossary/vbe-glossary.md#procedure), which sets a reference to an [object](../../Glossary/vbe-glossary.md#object).

## Syntax

[ **Public** |**Private** |**Friend** ] [ **Static** ] **Property** **Set**_name_**(** [ _arglist_**,** ] _reference_**)**
[ _statements_ ]
[ **Exit Property** ]
[ _statements_ ]

 **End Property**
The  **Property Set** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
|**Optional**|Optional. Indicates that the argument may or may not be supplied by the caller.|
|**Public**|Optional. Indicates that the  **Property** **Set** procedure is accessible to all other procedures in all[modules](../../Glossary/vbe-glossary.md#module). If used in a module that contains an  **Option Private** statement, the procedure is not available outside the[project](../../Glossary/vbe-glossary.md#project).|
|**Private**|Optional. Indicates that the  **Property** **Set** procedure is accessible only to other procedures in the module where it is declared.|
|**Friend**|Optional. Used only in a [class module](../../Glossary/vbe-glossary.md#class-module). Indicates that the  **Property Set** procedure is visible throughout the[project](../../Glossary/vbe-glossary.md#project), but not visible to a controller of an instance of an object.|
|**Static**|Optional. Indicates that the  **Property** **Set** procedure's local[variables](../../Glossary/vbe-glossary.md#variable) are preserved between calls. The **Static** attribute doesn't affect variables that are declared outside the **Property Set** procedure, even if they are used in the procedure.|
| _name_|Required. Name of the  **Property** **Set** procedure; follows standard variable naming conventions, except that the name can be the same as a **Property** **Get** or **Property Let** procedure in the same module.|
| _arglist_|Required. List of variables representing arguments that are passed to the  **Property** **Set** procedure when it is called. Multiple arguments are separated by commas.|
| _reference_|Required. Variable containing the object reference used on the right side of the object reference assignment.|
| _statements_|Optional. Any group of statements to be executed within the body of the  **Property** procedure.|

The  _arglist_ argument has the following syntax and parts:
[ **Optional** ] [ **ByVal** |**ByRef** ] [ **ParamArray** ] _varname_ [ **( )** ] [ **As**_type_ ] [ **=**_defaultvalue_ ]


|**Part**|**Description**|
|:-----|:-----|
|**Optional**|Optional. Indicates that an argument is not required. If used, all subsequent arguments in  _arglist_ must also be optional and declared using the **Optional** keyword. Note that it is not possible for the right side of a **Property Set**[expression](../../Glossary/vbe-glossary.md#expression) to be **Optional**.|
|**ByVal**|Optional. Indicates that the argument is passed [by value](../../Glossary/vbe-glossary.md#by-value).|
|**ByRef**|Optional. Indicates that the argument is passed [by reference](../../Glossary/vbe-glossary.md#by-reference).  **ByRef** is the default in Visual Basic.|
|**ParamArray**|Optional. Used only as the last argument in  _arglist_ to indicate that the final argument is an **Optional** array of **Variant** elements. The **ParamArray** keyword allows you to provide an arbitrary number of arguments. It may not be used with **ByVal**, **ByRef**, or **Optional**.|
| _varname_|Required. Name of the variable representing the argument; follows standard variable naming conventions.|
| _type_|Optional. [Data type](../../Glossary/vbe-glossary.md#Data-type) of the argument passed to the procedure; may be[Byte](../../Glossary/vbe-glossary.md#Byte), [Boolean](../../Glossary/vbe-glossary.md#Boolean), [Integer](../../Glossary/vbe-glossary.md#Integer), [Long](../../Glossary/vbe-glossary.md#Long), [Currency](../../Glossary/vbe-glossary.md#Currency), [Single](../../Glossary/vbe-glossary.md#Single), [Double](../../Glossary/vbe-glossary.md#Double), [Decimal](../../Glossary/vbe-glossary.md#Decimal) (not currently supported),[Date](../../Glossary/vbe-glossary.md#Date), [String](../../Glossary/vbe-glossary.md#String) (variable length only),[Object](../../Glossary/vbe-glossary.md#Object), [Variant](../../Glossary/vbe-glossary.md#Variant), or a specific [object type](../../Glossary/vbe-glossary.md#object-type). If the parameter is not  **Optional**, a[user-defined type](../../Glossary/vbe-glossary.md#user-defined-type) may also be specified.|
| _defaultvalue_|Optional. Any [constant](../../Glossary/vbe-glossary.md#constant) or constant expression. Valid for **Optional** parameters only. If the type is an **Object**, an explicit default value can only be **Nothing**.|

 **Note**  Every  **Property Set** statement must define at least one argument for the procedure it defines. That argument (or the last argument if there is more than one) contains the actual object reference for the property when the procedure defined by the **Property Set** statement is invoked. It is referred to as _reference_ in the preceding syntax. It can't be **Optional**.

## Remarks

If not explicitly specified using  **Public**, **Private**, or **Friend**, **Property** procedures are public by default. If **Static** isn't used, the value of local variables is not preserved between calls. The **Friend** keyword can only be used in class modules. However, **Friend** procedures can be accessed by procedures in any module of a project. A **Friend** procedure doesn't appear in the[type library](../../Glossary/vbe-glossary.md#type-library) of its parent class, nor can a **Friend** procedure be late bound.
All executable code must be in procedures. You can't define a  **Property** **Set** procedure inside another **Property**, **Sub**, or **Function** procedure.
The  **Exit Property** statement causes an immediate exit from a **Property** **Set** procedure. Program execution continues with the statement following the statement that called the **Property** **Set** procedure. Any number of **Exit Property** statements can appear anywhere in a **Property** **Set** procedure.
Like a  **Function** and **Property Get** procedure, a **Property Set** procedure is a separate procedure that can take arguments, perform a series of statements, and change the value of its arguments. However, unlike a **Function** and **Property Get** procedure, both of which return a value, you can only use a **Property Set** procedure on the left side of an object reference assignment (**Set** statement).

## Example

This example uses the  **Property Set** statement to define a property procedure that sets a reference to an object.


```vb
' The Pen property may be set to different Pen implementations. 
Property Set Pen(P As Object) 
 Set CurrentPen = P ' Assign Pen to object. 
End Property
```


