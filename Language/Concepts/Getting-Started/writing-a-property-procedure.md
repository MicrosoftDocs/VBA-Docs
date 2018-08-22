---
title: Writing a Property Procedure
keywords: vbcn6.chm1101346
f1_keywords:
- vbcn6.chm1101346
ms.prod: office
ms.assetid: 7ec62de7-4628-423e-54af-a49b0aef9d3c
ms.date: 06/08/2017
---


# Writing a Property Procedure

A  **Property** procedure is a series of Visual Basic[statements](../../Glossary/vbe-glossary.md#statement) that allow a programmer to create and manipulate custom properties.



-  **Property** procedures can be used to create read-only properties for[forms](../../Glossary/vbe-glossary.md#form), [standard modules](../../Glossary/vbe-glossary.md#standard-module), and [class modules](../../Glossary/vbe-glossary.md#class-module).
    
-  **Property** procedures should be used instead of **Public** variables in code that must be executed when the property value is set.
    
- Unlike  **Public** variables, **Property** procedures can have Help strings assigned to them in the[Object Browser](../../Glossary/vbe-glossary.md#Object-Browser).
    

When you create a  **Property** procedure, it becomes a property of the[module](../../Glossary/vbe-glossary.md#module) containing the procedure. Visual Basic provides the following three types of **Property** procedures:


|**Procedure**|**Description**|
|:-----|:-----|
|**Property Let**|A procedure that sets the value of a property.|
|**Property Get**|A procedure that returns the value a property.|
|**Property Set**|A procedure that sets a reference to an object.|

The syntax for declaring a  **Property** procedure is:
[ **Public** |**Private** ] [ **Static** ] **Property** { **Get** |**Let** |**Set** } _propertyname__ [( _arguments_ )] [ **As**_type_ ]
 _statements_
 **End Property**
 **Property** procedures are usually used in pairs: **Property Let** with **Property Get** and **Property Set** with **Property Get**. Declaring a **Property Get** procedure alone is like declaring a read-only property. Using all three **Property** procedure types together is only useful for **Variant** variables, since only a **Variant** can contain either an object or other data type information. **Property Set** is intended for use with objects; **Property Let** isn't.
The required arguments in  **Property** procedure declarations are shown in the following table:


|**Procedure**|**Declaration Syntax**|
|:-----|:-----|
|**Property Get**|**Property Get**_propname_ (1, …, _n_ ) **As**_type_|
|**Property Let**|**Property Let**_propname_ (1, …,,,, _n_, _n_ +1)|
|**Property Set**|**Property Set**_propname_ (1, …, _n_, _n_ +1)|

The first argument through the next to last argument (1, …,  _n_ ) must share the same names and data types in all **Property** procedures with the same name.
A  **Property Get** procedure declaration takes one less argument than the related **Property Let** and **Property Set** declarations. The data type of the **Property Get** procedure must be the same as the data type as the data type of the last argument ( _n_ +1) in the related **Property Let** and **Property Set** declarations. For example, if you declare the following **Property Let** procedure, the **Property Get** declaration must use arguments with the same name and data type as the arguments in the **Property Let** procedure.



```vb
Property Let Names(intX As Integer, intY As Integer, varZ As Variant) 
 ' Statement here. 
End Property 
 
Property Get Names(intX As Integer, intY As Integer) As Variant 
 ' Statement here. 
End Property 

```

The data type of the final argument in a  **Property Set** declaration must be either an[object type](../../Glossary/vbe-glossary.md#object-type) or a **Variant**.

