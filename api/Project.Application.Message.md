---
title: Application.Message method (Project)
keywords: vbapj.chm2
f1_keywords:
- vbapj.chm2
ms.prod: project-server
api_name:
- Project.Application.Message
ms.assetid: d601b101-5338-f404-e63e-6d1ce926a3d7
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Message method (Project)

Displays a message in a message box.


## Syntax

_expression_. `Message`( `_Message_`, `_Type_`, `_YesText_`, `_NoText_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Message_|Required|**String**|The message to display in the dialog box.|
| _Type_|Optional|**Long**|The buttons to include in the message dialog box. Can be one of the **[PjMessageType](Project.PjMessageType.md)** constants. The default value is **pjOKOnly**.|
| _YesText_|Optional|**String**|The text to be displayed on the **Yes** button. The YesText argument is ignored unless Type is **pjYesNo** or **pjYesNoCancel**. The default value is "Yes".|
| _NoText_|Optional|**String**|The text to be displayed on the **No** button. The NoText argument is ignored unless Type is **pjYesNo** or **pjYesNoCancel**. The default value is "No".|

## Return value

 **Boolean**


## Remarks

The **Message** method provides compatibility with the macro language used in Microsoft Project version 3. _x_. The **MsgBox** method in the VBA library should be used in new macros.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]