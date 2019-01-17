---
title: AllFunctions.Application property (Access)
keywords: vbaac10.chm12678
f1_keywords:
- vbaac10.chm12678
ms.prod: access
api_name:
- Access.AllFunctions.Application
ms.assetid: a71106f9-2949-c514-62aa-3c8cbff9cf09
ms.date: 06/08/2017
localization_priority: Normal
---


# AllFunctions.Application property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](Access.Application.md)** object and its related properties. Read-only **Application** object.


## Syntax

_expression_. `Application`

_expression_ A variable that represents an [AllFunctions](Access.AllFunctions.md) object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


[AllFunctions Collection](Access.AllFunctions.md)

