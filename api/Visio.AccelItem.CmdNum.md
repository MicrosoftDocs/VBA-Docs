---
title: AccelItem.CmdNum property (Visio)
keywords: vis_sdr.chm14513255
f1_keywords:
- vis_sdr.chm14513255
ms.prod: visio
api_name:
- Visio.AccelItem.CmdNum
ms.assetid: fb12e22d-671d-1f40-475c-714599fe0e37
ms.date: 06/24/2019
localization_priority: Normal
---


# AccelItem.CmdNum property (Visio)

Gets or sets the command ID associated with an accelerator. Read/write.


## Syntax

_expression_.**CmdNum**

_expression_ A variable that represents an **[AccelItem](Visio.AccelItem.md)** object.


## Return value

Integer


## Remarks

The **CmdNum** property should never be zero for an **AccelItem** object.

Valid command IDs are declared by the Visio type library in **[VisUICmds](Visio.visuicmds.md)**. They have the prefix **visCmd**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]