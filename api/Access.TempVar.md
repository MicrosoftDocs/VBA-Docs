---
title: TempVar Object (Access)
keywords: vbaac10.chm14063
f1_keywords:
- vbaac10.chm14063
ms.prod: access
api_name:
- Access.TempVar
ms.assetid: 4a0429e6-bcfa-7a8b-7030-6e88c2f1a71d
ms.date: 06/08/2017
---


# TempVar Object (Access)

Represents a variable that be used in Visual Basic for Applications (VBA) code or from a macro. 


## Remarks

A  **TempVar** objects provide a convenient way to exchange data between VBA procedures and macros.

Although a  **TempVar** object can be used to store information for use in VBA procedures, it does not have the same funcitonality as a VBA variable.


- By default, a  **TempVar** object remains in memory until Access is closed. You can use the **[Remove](./Access.TempVars.Remove.md)** method or the[RemoveTempVar](http://msdn.microsoft.com/library/7bcc5010-3e30-ecef-2c5d-a35e73c8e325%28Office.15%29.aspx) macro action to remove a **TempVar** object.
    
- In VBA, a  **TempVar** object is accessible only to the members of the Access **[Application](./Access.Application.md)** object, referenced databases, or add-ins.
    
- A  **TempVar** object can store only text or numeric data. **TempVar** objects cannot store objects.
    
To refer to a  **TempVar** object in a collection by its ordinal number or by its **Name** property setting, use the following syntax form:


-  **TempVar** ![name]
    

|**Name**|
|:-----|
|[Name](./Access.TempVar.Name.md)|
|[Value](./Access.TempVar.Value.md)|

## See also


[Access Object Model Reference](./overview/object-model-access-vba-reference.md)
[TempVar Object Members](http://msdn.microsoft.com/library/1d8ac3a8-3116-6ce5-90c0-83265d7b79c4%28Office.15%29.aspx)
