---
title: ReturnVar object (Access)
keywords: vbaac10.chm14689
f1_keywords:
- vbaac10.chm14689
ms.prod: access
api_name:
- Access.ReturnVar
ms.assetid: 8ad5254d-a249-46ba-ac5d-14943179ce05
ms.date: 03/21/2019
localization_priority: Normal
---


# ReturnVar object (Access)

Represents a variable that was initialized by the **SetReturnVar** function in a Data Macro.


## Remarks

A **ReturnVar** object provides a convenient way to use values set in a Data Macro.

Although a **ReturnVar** object can be used to store information for use in VBA procedures, it does not have the same functionality as a VBA variable.

By default, a **ReturnVar** object remains in memory until the next time the **[RunDataMacro](Access.DoCmd.RunDataMacro.md)** method is used.
    
A **ReturnVar** object can store only text or numeric data. **ReturnVar** objects cannot store objects.
    
To refer to a **TempVar** object in a collection by its ordinal number or by its **Name** property setting, use the following syntax form.

```vb
ReturnVars![name] 

```


## Properties

- [Name](Access.ReturnVar.Name.md)
- [Value](Access.ReturnVar.Value.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]