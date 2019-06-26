---
title: Application.MakeLocalCalendarEnterprise method (Project)
keywords: vbapj.chm2369
f1_keywords:
- vbapj.chm2369
ms.prod: project-server
api_name:
- Project.Application.MakeLocalCalendarEnterprise
ms.assetid: deb355ad-39ca-77cd-7d0d-f5915c7185da
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.MakeLocalCalendarEnterprise method (Project)

Converts a local calendar to an enterprise calendar.


## Syntax

_expression_. `MakeLocalCalendarEnterprise`( `_OldName_`, `_NewName_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _OldName_|Optional|**String**|Name of the local calendar.|
| _NewName_|Optional|**String**|Name of the Enterprise calendar.|

## Return value

 **Boolean**


## Remarks

The  _NewName_ parameter is not used. For example, if a local calendar is named "TestCal" and you execute the code `MakeLocalCalendarEnterprise OldName:="TestCal", NewName:="New TestCal"` , the result is an enterprise calendar named "TestCal".

To create a local calendar when Project Professional is logged on to Project Server, you must check  **Allow projects to use local base calendars** on the Additional Server Settings page in Project Web Access. Restart Project Professional after changing the setting in Project Web Access.


## Example

The following example creates a local calendar named  **TestCal**, and then saves it as an enterprise calendar with the same name. If Project Professional is not logged on Project Server, **MakeLocalCalendarEnterprise** results in a run-time error 1100.


```vb
Sub TestCalendar() 
 BaseCalendarCreate Name:="TestCal" 
 MakeLocalCalendarEnterprise OldName:="TestCal" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]