---
title: Application.Run method (Excel)
keywords: vbaxl10.chm132104
f1_keywords:
- vbaxl10.chm132104
ms.prod: excel
api_name:
- Excel.Application.Run
ms.assetid: 3e0167ab-b101-018f-0f89-ada116b8bb72
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.Run method (Excel)

Runs a macro or calls a function. This can be used to run a macro written in Visual Basic or the Microsoft Excel macro language, or to run a function in a DLL or XLL.


## Syntax

_expression_.**Run** (_Macro_, _Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_, _Arg6_, _Arg7_, _Arg8_, _Arg9_, _Arg10_, _Arg11_, _Arg12_, _Arg13_, _Arg14_, _Arg15_, _Arg16_, _Arg17_, _Arg18_, _Arg19_, _Arg20_, _Arg21_, _Arg22_, _Arg23_, _Arg24_, _Arg25_, _Arg26_, _Arg27_, _Arg28_, _Arg29_, _Arg30_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Macro_|Optional| **Variant**|The macro to run.<br/><br/>This can be either a string with the macro name, a **[Range](Excel.Range(object).md)** object indicating where the function is, or a register ID for a registered DLL (XLL) function.<br/><br/>If a string is used, the string will be evaluated in the context of the active sheet.|
| _Arg1_&ndash;_Arg30_|Optional| **Variant**|An argument that should be passed to the function.|

## Return value

Variant


## Remarks

You cannot use named arguments with this method. Arguments must be passed by position.

The **Run** method returns whatever the called macro returns.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
