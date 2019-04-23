---
title: Application.HinstancePtr property (Excel)
keywords: vbaxl10.chm133334
f1_keywords:
- vbaxl10.chm133334
ms.prod: excel
api_name:
- Excel.Application.HinstancePtr
ms.assetid: fddc40e9-08fc-34ef-60b2-41e8afa86575
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.HinstancePtr property (Excel)

Returns a handle to the instance of Excel represented by the specified **Application** object. Read-only **Variant**.


## Syntax

_expression_.**HinstancePtr**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

This property returns a correct handle in both the 32-bit and 64-bit versions of Excel. It extends the functionality of the **[Hinstance](Excel.Application.Hinstance.md)** property of the **Application** object, which only works correctly in the 32-bit version of Excel.

The ideal data type to use with this property is the **[LongPtr](../language/reference/User-Interface-Help/longptr-data-type.md)** data type. Assigning the value returned by this property to a **LongPtr** variable will work as expected in both 32-bit and 64-bit versions of Excel. The property is defined as **Variant** for internal implementation reasons. However, it always returns a 32-bit value on 32-bit systems and a 64-bit value on 64-bit systems.

This property only works starting with Excel, and is only required with the 64-bit version of Excel. If you must write code that will also work with earlier versions of Excel, in order to avoid compilation errors, read this property under an `#if Win64` conditional compilation directive, and use the **Hinstance** property under the `#else` directive.

Note that this property works fine in both 32-bit and 64-bit environments starting with Excel. Therefore, if your code is intended to be used only with Excel or later, either 32-bit or 64-bit, it can read this property without conditional compilation.

For more information about how to use VBA in 64-bit environments, see [64-bit Visual Basic for Applications overview](../Language/Concepts/Getting-Started/64-bit-visual-basic-for-applications-overview.md).


## Example

In this example, a message box displays the Excel instance handle to the user.

```vb
Sub CheckHinstance() 
    MsgBox Application.HinstancePtr 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]