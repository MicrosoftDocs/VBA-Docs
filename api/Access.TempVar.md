---
title: TempVar object (Access)
keywords: vbaac10.chm14063
f1_keywords:
- vbaac10.chm14063
ms.prod: access
api_name:
- Access.TempVar
ms.assetid: 4a0429e6-bcfa-7a8b-7030-6e88c2f1a71d
ms.date: 03/21/2019
localization_priority: Normal
---


# TempVar object (Access)

Represents a variable that can be used in Visual Basic for Applications (VBA) code or from a macro. 


## Remarks

**TempVar** objects provide a convenient way to exchange data between VBA procedures and macros.

Although a **TempVar** object can be used to store information for use in VBA procedures, it does not have the same functionality as a VBA variable.

- By default, a **TempVar** object remains in memory until Access is closed. You can use the **[Remove](Access.TempVars.Remove.md)** method or the [RemoveTempVar](overview/Access.md) macro action to remove a **TempVar** object.
    
- In VBA, a **TempVar** object is accessible only to the members of the Access **[Application](Access.Application.md)** object, referenced databases, or add-ins.
    
- A **TempVar** object can store only text or numeric data. **TempVar** objects cannot store objects.
    
To refer to a **TempVar** object in a collection by its ordinal number or by its **Name** property setting, use the following syntax form:

- **TempVar**![name]

## Properties   

- [Name](Access.TempVar.Name.md)
- [Value](Access.TempVar.Value.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
