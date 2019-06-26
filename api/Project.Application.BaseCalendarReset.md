---
title: Application.BaseCalendarReset method (Project)
keywords: vbapj.chm617
f1_keywords:
- vbapj.chm617
ms.prod: project-server
api_name:
- Project.Application.BaseCalendarReset
ms.assetid: 43c842b2-146b-f080-f88b-c1e0ef5526d8
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.BaseCalendarReset method (Project)

Resets a base calendar.


## Syntax

_expression_. `BaseCalendarReset`( `_Name_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|**String**. The name of the base calendar to reset.|

## Return value

 **Boolean**


## Remarks

Base calendars have the following default characteristics:




- Monday through Friday are working days with two shifts (8 A.M. to 12 P.M. and 1 P.M. to 5 P.M.).
    
- Saturday and Sunday are nonworking days.
    



## Example

The following example resets the Standard base calendar to the default settings.


```vb
Sub RestoreBaseCalendar() 
 BaseCalendarReset Name:="Standard" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]