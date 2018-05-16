---
title: ObjectFrame.Application Property (Access)
keywords: vbaac10.chm11547
f1_keywords:
- vbaac10.chm11547
ms.prod: access
api_name:
- Access.ObjectFrame.Application
ms.assetid: 29e3d68f-4f67-793d-6976-e18e290145fe
ms.date: 06/08/2017
---


# ObjectFrame.Application Property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](Access.Application.md)** object and its related properties. Read-only **Application** object.


## Syntax

 _expression_. **Application**

 _expression_ A variable that represents an **ObjectFrame** object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


#### Concepts


[ObjectFrame Object](Access.ObjectFrame.md)

