---
title: MacroError object (Access)
keywords: vbaac10.chm14053
f1_keywords:
- vbaac10.chm14053
ms.prod: access
api_name:
- Access.MacroError
ms.assetid: 556c4fdb-c88e-a102-bccd-71bd53c9cffb
ms.date: 03/21/2019
localization_priority: Normal
---


# MacroError object (Access)

Represents the properties of a run-time error that occurs in a macro.


## Remarks

When an error occurs in a macro, information about the error is stored in the **MacroError** object. If you have not used the OnError action to suppress error messages, the macro stops and the error information is displayed in a standard error message. However, if you have used the OnError action to suppress error messages, you may want to use the information stored in the **MacroError** object in a condition or a custom error message.

After an error has been handled, the information in the **MacroError** object is out of date, so it is a good idea to clear the object by using the ClearMacroError action. This resets the error number in the **MacroError** object back to zero, and clears any other information about the error that is stored in the object, such as the error description, macro name, action name, condition, and arguments. This way, you can inspect the **MacroError** object again later to see if another error has occurred.

The **MacroError** object contains information about only one error at a time. If more than one error has occurred in a macro, the **MacroError** object contains information about only the last one.

The **MacroError** object does not contain information about run-time errors that occur when running Visual Basic for Applications (VBA) code. For more information about handling run-time errors in VBA, see [Elements of run-time error handling](../access/Concepts/Error-Codes/elements-of-run-time-error-handling.md).


## Properties

- [ActionName](Access.MacroError.ActionName.md)
- [Arguments](Access.MacroError.Arguments.md)
- [Condition](Access.MacroError.Condition.md)
- [Description](Access.MacroError.Description.md)
- [MacroName](Access.MacroError.MacroName.md)
- [Number](Access.MacroError.Number.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]