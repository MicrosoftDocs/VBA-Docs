---
title: Application.DDETerminateAll method (Access)
keywords: vbaac10.chm12544
f1_keywords:
- vbaac10.chm12544
ms.prod: access
api_name:
- Access.Application.DDETerminateAll
ms.assetid: 0d2a5e65-c10a-1e78-a0a3-573b9ed804d4
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.DDETerminateAll method (Access)

You can use the **DDETerminateAll** statement to close all open dynamic data exchange (DDE) channels.


## Syntax

_expression_.**DDETerminateAll**

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Return value

Nothing


## Remarks

For example, suppose you've opened two DDE channels between Microsoft Excel and Microsoft Access, one to retrieve system information about Excel and one to transfer data. You can use the **DDETerminateAll** statement to close both channels simultaneously.

Using the **DDETerminateAll** statement is equivalent to executing a **[DDETerminate](Access.Application.DDETerminate.md)** statement for each open channel number. Like the **DDETerminate** statement, the **DDETerminateAll** statement has no effect on active DDE link expressions in fields on forms or reports.

If there are no DDE channels open, the **DDETerminateAll** statement runs without causing a run-time error.

> [!TIP] 
> - If you interrupt a procedure that performs DDE, you may inadvertently leave channels open. To avoid exhausting system resources, use the **DDETerminateAll** statement in your code or from the Immediate (lower) pane of the Debug window while debugging code that performs DDE.
> - If you need to manipulate another application's objects from Access, you may want to consider using Automation.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]