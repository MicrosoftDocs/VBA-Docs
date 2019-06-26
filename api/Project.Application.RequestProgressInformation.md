---
title: Application.RequestProgressInformation method (Project)
ms.prod: project-server
api_name:
- Project.Application.RequestProgressInformation
ms.assetid: a86ec09d-f9c8-07e3-68f4-898c604c3600
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.RequestProgressInformation method (Project)

Requests progress information from resources, republishes, and saves the active project. .


## Syntax

_expression_. `RequestProgressInformation`( `_ShowDialog_`, `_ItemsScope_`, `_NotifyTaskLead_`, `_NotificationText_`, `_ReportingPeriodFrom_`, `_ReportingPeriodTo_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShowDialog_|Optional|**Boolean**|**True** if Project displays the corresponding dialog box for the message type. The default value is **False**.|
| _ItemsScope_|Optional|**Long**|Specifies the scope of assignments to be published. Can be one of the following  **[PjPublishScope](Project.PjPublishScope.md)** constants: **pjPublishScopeAll**, **pjPublishScopeDefault**, **pjPublishScopeSelected**, or **pjPublishScopeVisible**. The default value is **pjPublishScopeAll**.|
| _NotifyTaskLead_|Optional|**Boolean**|**True** if Project only notifies the task lead for delegated tasks with a lead. The default value is **False**.|
| _NotificationText_|Optional|**String**|The body text for the email notification.|
| _ReportingPeriodFrom_|Optional|**Variant**|The beginning date of the reporting period for assignment status. This affects the user's filtered tasks view or MAPI message.|
| _ReportingPeriodTo_|Optional|**Variant**|The end date of the reporting period for assignment status. This affects the users filtered tasks view or MAPI message.|

## Remarks

Using the  **RequestProgressInformation** method with no arguments displays the **Request Progress Information** dialog box. The **RequestProgressInformation** method is available only in Project Professional.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]