---
title: TempVars.Application Property (Access)
keywords: vbaac10.chm14065
f1_keywords:
- vbaac10.chm14065
ms.prod: access
api_name:
- Access.TempVars.Application
ms.assetid: 250a64f6-d0a2-d816-1211-c56d90de0e70
ms.date: 06/08/2017
---


# TempVars.Application Property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](Access.Application.md)** object and its related properties. Read-only **Application** object.


## Syntax

 _expression_. **Application**

 _expression_ A variable that represents a **TempVars** object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


#### Concepts


[TempVars Collection](Access.TempVars.md)

