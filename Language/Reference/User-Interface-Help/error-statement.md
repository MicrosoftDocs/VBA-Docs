---
title: Error statement (VBA)
keywords: vblr6.chm1008913
f1_keywords:
- vblr6.chm1008913
ms.prod: office
ms.assetid: b657920d-b28c-0c6b-8020-9d37e9f10f6c
ms.date: 12/03/2018
localization_priority: Normal
---


# Error statement

Simulates the occurrence of an error.

## Syntax

**Error** _errornumber_

The required _errornumber_ can be any valid [error number](../../Glossary/vbe-glossary.md#error-number).

## Remarks

The **Error** statement is supported for backward compatibility. In new code, especially when creating objects, use the **[Err](err-object.md)** object's **[Raise](raise-method.md)** method to generate [run-time errors](../../Glossary/vbe-glossary.md#run-time-error).

If _errornumber_ is defined, the **Error** statement calls the error handler after the [properties](../../Glossary/vbe-glossary.md#property) of the **Err** object are assigned the following default values:


|Property|Value|
|:-----|:-----|
|**Number**|Value specified as [argument](../../Glossary/vbe-glossary.md#argument) to **Error** statement. Can be any valid error number.|
|**Source**|Name of the current Visual Basic [project](../../Glossary/vbe-glossary.md#project).|
|**Description**|[String expression](../../Glossary/vbe-glossary.md#string-expression) corresponding to the return value of the **Error** function for the specified **Number**, if this string exists. If the string doesn't exist, **Description** contains a zero-length string ("").|
|**HelpFile**|The fully qualified drive, path, and file name of the appropriate Visual Basic Help file.|
|**HelpContext**|The appropriate Visual Basic Help file context ID for the error corresponding to the **Number** property.|
|**LastDLLError**|Zero.|

If no error handler exists or if none is enabled, an error message is created and displayed from the **Err** object properties.

> [!NOTE] 
> Not all Visual Basic [host applications](../../Glossary/vbe-glossary.md#host-application) can create objects; for example, hosts running versions of Visual Basic for Applications earlier than 4.0 cannot create objects. Because **Err** is a function returning an **ErrObject** instance, it cannot be used in these early versions. To know what version of VBA your host application is running, see the **About** information for your Visual Basic Editor (VBE), and see your host application's documentation to determine whether it can create [classes](../../Glossary/vbe-glossary.md#class) and objects.

## Example

This example uses the **Error** statement to simulate error number 11.

```vb
On Error Resume Next ' Defer error handling. 
Error 11 ' Simulate the "Division by zero" error. 

```


## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]