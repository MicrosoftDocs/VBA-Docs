---
title: AllQueries.Application property (Access)
keywords: vbaac10.chm12678
f1_keywords:
- vbaac10.chm12678
ms.prod: access
api_name:
- Access.AllQueries.Application
ms.assetid: a123a13e-57f7-a637-e533-c5a6da3f3f94
ms.date: 06/08/2017
---


# AllQueries.Application property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](Access.Application.md)** object and its related properties. Read-only **Application** object.


## Syntax

_expression_. `Application`

_expression_ A variable that represents an [AllQueries](Access.AllQueries.md) object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


[AllQueries Collection](Access.AllQueries.md)

