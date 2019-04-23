---
title: Call statement (VBA)
keywords: vblr6.chm1008863
f1_keywords:
- vblr6.chm1008863
ms.prod: office
ms.assetid: 6232c5cd-8bfe-2316-a0f6-6323db933357
ms.date: 12/03/2018
localization_priority: Normal
---


# Call statement

Transfers control to a **[Sub](sub-statement.md)** [procedure](../../Glossary/vbe-glossary.md#procedure), **[Function](function-statement.md)** procedure, or [dynamic-link library (DLL)](../../Glossary/vbe-glossary.md#dynamic-link-library-dll) procedure.

## Syntax

[ **Call** ] _name_ [ _argumentlist_ ]

<br/>

The **Call** statement syntax has these parts:

|Part|Description|
|:-----|:-----|
|**Call**|Optional; [keyword](../../Glossary/vbe-glossary.md#keyword). If specified, you must enclose _argumentlist_ in parentheses.<br/><br/>For example: `Call MyProc(0)` |
|_name_|Required. Name of the procedure to call.|
|_argumentlist_|Optional. Comma-delimited list of [variables](../../Glossary/vbe-glossary.md#variable), [arrays](../../Glossary/vbe-glossary.md#array), or [expressions](../../Glossary/vbe-glossary.md#expression) to pass to the procedure. Components of _argumentlist_ may include the keywords **ByVal** or **ByRef** to describe how the [arguments](../../Glossary/vbe-glossary.md#argument) are treated by the called procedure.<br/><br/>However, **ByVal** and **ByRef** can be used with **Call** only when calling a DLL procedure. On the Macintosh, **ByVal** and **ByRef** can be used with **Call** when making a call to a Macintosh code resource.|

## Remarks

You are not required to use the **Call** keyword when calling a procedure. However, if you use the **Call** keyword to call a procedure that requires arguments, _argumentlist_ must be enclosed in parentheses. If you omit the **Call** keyword, you also must omit the parentheses around _argumentlist_. If you use either **Call** syntax to call any intrinsic or user-defined function, the function's return value is discarded.

To pass a whole array to a procedure, use the array name followed by empty parentheses.

## Example

This example illustrates how the **Call** statement is used to transfer control to a **Sub** procedure, an intrinsic function, and a dynamic-link library (DLL) procedure. DLLs are not used on the Macintosh.


```vb
' Call a Sub procedure. 
Call PrintToDebugWindow("Hello World")     
' The above statement causes control to be passed to the following 
' Sub procedure. 
Sub PrintToDebugWindow(AnyString) 
    Debug.Print AnyString    ' Print to the Immediate window. 
End Sub 
 
' Call an intrinsic function. The return value of the function is 
' discarded. 
Call Shell(AppName, 1)    ' AppName contains the path of the  
        ' executable file. 
 
' Call a Microsoft Windows DLL procedure. The Declare statement must be  
' Private in a Class Module, but not in a standard Module. 
Private Declare Sub MessageBeep Lib "User" (ByVal N As Integer) 
Sub CallMyDll() 
    Call MessageBeep(0)    ' Call Windows DLL procedure. 
    MessageBeep 0    ' Call again without Call keyword. 
End Sub
```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
