---
title: MacroError.Condition property (Access)
keywords: vbaac10.chm14046
f1_keywords:
- vbaac10.chm14046
ms.prod: access
api_name:
- Access.MacroError.Condition
ms.assetid: 4210aff0-6f94-9b09-3a03-bbfb2f2b2494
ms.date: 03/22/2019
localization_priority: Normal
---


# MacroError.Condition property (Access)

Gets the condition of the macro action that was executing when an error occurred. Read-only **String**.


## Syntax

_expression_.**Condition**

_expression_ A variable that represents a **[MacroError](Access.MacroError.md)** object.


## Remarks

When an error occurs in a macro, information about the error is stored in the **MacroError** object. If you have not used the OnError action to suppress error messages, the macro stops and the error information is displayed in a standard error message. However, if you have used the OnError action to suppress error messages, you may want to use the information stored in the **MacroError** object in a condition or a custom error message.

After an error has been handled, the information in the **MacroError** object is out of date, so it is a good idea to clear the object by using the ClearMacroError action. This resets the error number in the **MacroError** object back to zero, and clears any other information about the error that is stored in the object, such as the error description, macro name, action name, condition, and arguments. This way, you can inspect the **MacroError** object again later to see if another error has occurred.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]