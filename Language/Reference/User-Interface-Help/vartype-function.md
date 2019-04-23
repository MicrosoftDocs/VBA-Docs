---
title: VarType function (Visual Basic for Applications)
keywords: vblr6.chm1009057
f1_keywords:
- vblr6.chm1009057
ms.prod: office
ms.assetid: 7422fba5-7ea9-1d91-fc0e-5694c352d2d0
ms.date: 04/17/2019
localization_priority: Normal
---


# VarType function

Returns an **Integer** indicating the subtype of a [variable](../../Glossary/vbe-glossary.md#variable), or the type of an object's default [property](../../Glossary/vbe-glossary.md#property).

## Syntax

**VarType**(_varname_)

The required _varname_ [argument](../../Glossary/vbe-glossary.md#argument) is a [Variant](../../Glossary/vbe-glossary.md#variant-data-type) containing any variable except a variable of a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type).
 
## Return values

Either one of the following constants or the summation of a number of them is returned.

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbEmpty**|0|[Empty](../../Glossary/vbe-glossary.md#empty) (uninitialized)|
|**vbNull**|1|[Null](../../Glossary/vbe-glossary.md#null) (no valid data)|
|**vbInteger**|2|**Integer**|
|**vbLong**|3|Long integer|
|**vbSingle**|4|Single-precision floating-point number|
|**vbDouble**|5|Double-precision floating-point number|
|**vbCurrency**|6|Currency value|
|**vbDate**|7|Date value|
|**vbString**|8|**String**|
|**vbObject**|9|Object|
|**vbError**|10|Error value|
|**vbBoolean**|11|Boolean value|
|**vbVariant**|12|**Variant** (used only with [arrays](../../Glossary/vbe-glossary.md#array) of variants)|
|**vbDataObject**|13|A data access object|
|**vbDecimal**|14|Decimal value|
|**vbByte**|17|Byte value|
|**vbLongLong**|20|[LongLong](longlong-data-type.md) integer (valid on 64-bit platforms only)|
|**vbUserDefinedType**|36|Variants that contain user-defined types|
|**vbArray**|8192|Array (always added to another constant when returned by this function)|

> [!NOTE] 
> These [constants](../../Glossary/vbe-glossary.md#constant) are specified by Visual Basic for Applications. The names can be used anywhere in your code in place of the actual values.

## Remarks

If an object is passed and has a default property, **VarType**(_object_) returns the type of the object's default property.

The **VarType** function never returns the value for **vbArray** by itself. It is always added to some other value to indicate an array of a particular type. For example, the value returned for an array of integers is calculated as **vbInteger** + **vbArray**, or 8194. 

The constant **vbVariant** is only returned in conjunction with **vbArray** to indicate that the argument to the **VarType** function is an array of type **Variant**. 

## Example

This example uses the **VarType** function to determine the subtypes of different variables, and in one case, the type of an object's default property.

```vb
Dim MyCheck
Dim IntVar, StrVar, DateVar, AppVar, ArrayVar
' Initialize variables.
IntVar = 459: StrVar = "Hello World": DateVar = #2/12/1969#
Set AppVar = Excel.Application
ArrayVar = Array("1st Element", "2nd Element")
' Run VarType function on different types.
MyCheck = VarType(IntVar)   ' Returns 2.
MyCheck = VarType(DateVar)  ' Returns 7.
MyCheck = VarType(StrVar)   ' Returns 8.
MyCheck = VarType(AppVar)   ' Returns 8 (vbString)
                            ' even though AppVar is an object.
MyCheck = VarType(ArrayVar) ' Returns 8204 which is
                            ' `8192 + 12`, the computation of
                            ' `vbArray + vbVariant`.
```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
