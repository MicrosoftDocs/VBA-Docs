---
title: Shift.Index property (Project)
ms.prod: project-server
api_name:
- Project.Shift.Index
ms.assetid: dae37122-f745-2728-5004-b3b3d7ad188a
ms.date: 06/08/2017
localization_priority: Normal
---


# Shift.Index property (Project)

Gets the index of a  **Shift** object in the containing object. Read-only **Integer**.


## Syntax

_expression_.**Index**

_expression_ A variable that represents a [Shift](./Project.Shift.md) object.


## Remarks

Following are the objects that can contain  **Shift** objects:


-  **Day**
    
-  **Month**
    
-  **WeekDay**
    
-  **WorkWeekDay**
    
-  **Year**
    


 **Shift** objects are accessed using the **Shift1**... **Shift5** properties. Because Project defines five shifts, the **Index** property can have only the values 1 through 5.

The  **Index** properties of different objects are used in similar ways. For an example, see the **[Index](Project.Project.Index.md)** property of the **Project** object.


## Example

The following command in the Immediate window of the VBE prints the value 2.


```vb
? ActiveProject.Calendar.WeekDays.Item(3).Shift2.Index
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]