---
title: Exceptions.Add method (Project)
ms.prod: project-server
api_name:
- Project.Exceptions.Add
ms.assetid: a20cbcdf-d764-de46-d57f-0cc283665129
ms.date: 06/08/2017
localization_priority: Normal
---


# Exceptions.Add method (Project)

Adds an  **Exception** object to an **Exceptions** collection.


## Syntax

_expression_.**Add** (_Type_, _Start_, _Finish_, _Occurrences_, _Name_, _Period_, _DaysOfWeek_, _MonthPosition_, _MonthItem_, _Month_, _MonthDay_)

_expression_ A variable that represents an 'Exceptions' object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**Long**|Exception type. Can be one of the  **[PjExceptionType](Project.PjExceptionType.md)** constants.|
| _Start_|Required|**Variant**|Start date of the exception.|
| _Finish_|Optional|**Variant**|Finish date of the exception.|
| _Occurrences_|Optional|**Long**|Number of occurrences.|
| _Name_|Optional|**String**|Name of the  **Exception** object|
| _Period_|Optional|**Long**|Number for exception recurrence.|
| _DaysOfWeek_|Optional|**Long**|The days on which the exception occurs. Can be a combination of  **[PjWeekday](Project.PjWeekday.md)** constants.|
| _MonthPosition_|Optional|**Long**|Value for the  **Monthly** type exceptions. Can be one of the **[PjExceptionPosition](Project.PjExceptionPosition.md)** contants.|
| _MonthItem_|Optional|**Long**|Value for the  **Monthly** type exceptions. Can be one of the following **PjExceptionItem** constants: **pjSunday**, **pjMonday**, **pjTuesday**, **pjWednesday**, **pjThursday**, **pjFriday**, and **pjSaturday**.|
| _Month_|Optional|**Long**|Specifies the month, if the Type argument is  **pjYearlyMonthDay** or **pjYearlyPositional**. Can be one of the **[pjMonth](Project.PjMonth.md)** constants.|
| _MonthDay_|Optional|**Long**|Day of month for  **ByMonthDay** type exceptions.|

## Return value

 **Exception**


## See also


[Exceptions Collection Object](Project.exceptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]