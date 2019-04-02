---
title: PreviewPane.Session property (Outlook)
keywords: vbaol11.chm3636
f1_keywords:
- vbaol11.chm3636
ms.assetid: 54509e05-d255-b96e-f037-14282791ea55
ms.date: 06/08/2017
ms.prod: outlook
localization_priority: Normal
---


# PreviewPane.Session property (Outlook)

Returns the [NameSpace](Outlook.NameSpace.md) for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a 'PreviewPane' object.


## Remarks

The  **Session** property and the [GetNamespace](Outlook.Application.GetNamespace.md) method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:

 `Set objNamespace = Application.Getnamespace("MAPI")`

 `SetjobSession = Application.Session`


## See also


[PreviewPane object (Outlook)](Outlook.previewpane.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]