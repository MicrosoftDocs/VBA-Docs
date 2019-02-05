---
title: Application.DDEExecute method (Access)
keywords: vbaac10.chm12540
f1_keywords:
- vbaac10.chm12540
ms.prod: access
api_name:
- Access.Application.DDEExecute
ms.assetid: 9828607e-a2e3-15e2-699a-12fb2dc9e897
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.DDEExecute method (Access)

You can use the **DDEExecute** statement to send a command from a client application to a server application over an open dynamic data exchange (DDE) channel.


## Syntax

_expression_.**DDEExecute** (_ChanNum_, _Command_)

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ChanNum_|Required|**Variant**|A channel number, the long integer returned by the **[DDEInitiate](Access.Application.DDEInitiate.md)** function.|
| _Command_|Required|**String**|A command recognized by the server application. Check the server application's documentation for a list of these commands.|

## Remarks

For example, suppose you've opened a DDE channel in Microsoft Access to transfer text data from a Microsoft Excel spreadsheet into an Access database. Use the **DDEExecute** statement to send the **New** command to Excel to specify that you wish to open a new spreadsheet. In this example, Access acts as the client application, and Excel acts as the server application.

The value of the _command_ argument depends on the application and topic specified when the channel indicated by the _channum_ argument is opened. An error occurs if the _channum_ argument isn't an integer corresponding to an open channel or if the other application can't carry out the specified command.

From Visual Basic, you can use the **DDEExecute** statement only to send commands to another application. For information about sending commands to Access from another application, see [Use Microsoft Access as a DDE Server](overview/Access.md).

If you need to manipulate another application's objects from Microsoft Access, you may want to consider using Automation.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]