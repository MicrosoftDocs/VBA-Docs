---
title: MacroError.MacroName Property (Access)
keywords: vbaac10.chm14051
f1_keywords:
- vbaac10.chm14051
ms.prod: access
api_name:
- Access.MacroError.MacroName
ms.assetid: 9f36dde0-4b4f-67ef-0b84-765c8e848097
ms.date: 06/08/2017
---


# MacroError.MacroName Property (Access)

Gets the name of the macro that was running when an error occurred. Read-only  **String**.


## Syntax

 _expression_. **MacroName**

 _expression_ A variable that represents a **MacroError** object.


## Remarks

When an error occurs in a macro, information about the error is stored in the  **MacroError** object. If you have not used the **OnError** action to suppress error messages, the macro stops and the error information is displayed in a standard error message. However, if you have used the **OnError** action to suppress error messages, you may want to use the information stored in the **MacroError** object in a condition or a custom error message.

After an error has been handled, the information in the  **MacroError** object is out of date, so it is a good idea to clear the object using the **ClearMacroError** action. This resets the error number in the **MacroError** object back to zero, and clears any other information about the error that is stored in the object, such as the error description, macro name, action name, condition, and arguments. This way, you can inspect the **MacroError** object again later to see if another error has occurred.


## See also


#### Concepts


[MacroError Object](Access.MacroError.md)

