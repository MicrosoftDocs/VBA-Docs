---
title: AllModules.Application Property (Access)
keywords: vbaac10.chm12678
f1_keywords:
- vbaac10.chm12678
ms.prod: access
api_name:
- Access.AllModules.Application
ms.assetid: 9e6ea147-4e11-7f9e-b289-71b832630e13
ms.date: 06/08/2017
---


# AllModules.Application Property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](Access.Application.md)** object and its related properties. Read-only **Application** object.


## Syntax

 _expression_. **Application**

 _expression_ A variable that represents an **AllModules** object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


#### Concepts


[AllModules Collection](Access.AllModules.md)

