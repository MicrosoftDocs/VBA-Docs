---
title: Application.DrawingProperties method (Project)
keywords: vbapj.chm2307
f1_keywords:
- vbapj.chm2307
ms.prod: project-server
api_name:
- Project.Application.DrawingProperties
ms.assetid: 8d63be84-6321-c0b2-27f0-945baf349714
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DrawingProperties method (Project)

Displays the **Format Drawing** dialog box, which prompts the user to customize the active drawing object.


## Syntax

_expression_. `DrawingProperties`( `_SizePositionTab_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SizePositionTab_|Optional|**Boolean**|**True** if the **Size & Position** tab of the **Format Drawing** dialog box appears. **False** if the **Line & Fill** tab appears.|

## Return value

 **Boolean**


## Remarks

The **DrawingProperties** method displays an error unless a drawing object is active.

The **DrawingProperties** method has the same effect as the **Properties** command in the **Drawing** drop-down menu on the **Format** tab in the Ribbon.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]