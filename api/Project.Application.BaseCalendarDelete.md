---
title: Application.BaseCalendarDelete method (Project)
keywords: vbapj.chm619
f1_keywords:
- vbapj.chm619
ms.service: project-server
api_name:
- Project.Application.BaseCalendarDelete
ms.assetid: f9583bd7-6ddb-7115-b7ca-c0e4e8b033e1
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Application.BaseCalendarDelete method (Project)

Deletes a base calendar.


## Syntax

_expression_. `BaseCalendarDelete`( `_Name_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|**String**. The name of the base calendar to delete.|

## Return value

 **Boolean**


## Example

The following example deletes the base calendar entered by the user.


```vb
Sub DeleteCalendar() 
 
 Dim CalendarName As String 
 
 CalendarName = InputBox$("Enter name of base calendar to delete:") 
 BaseCalendarDelete Name:=CalendarName 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]