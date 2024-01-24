---
title: DoCmd.Hourglass method (Access)
keywords: vbaac10.chm4155
f1_keywords:
- vbaac10.chm4155
api_name:
- Access.DoCmd.Hourglass
ms.assetid: e032e879-6ce4-982d-08cb-f9622c000b11
ms.date: 09/07/2021
ms.localizationpriority: medium
---

# DoCmd.Hourglass method (Access)

The **Hourglass** method carries out the Hourglass action in Visual Basic.


## Syntax

_expression_.**Hourglass** (_HourglassOn_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _HourglassOn_|Required|**Variant**|Use **True** (1) to display the hourglass icon (or another icon you've chosen). Use **False** (0) to display the normal mouse pointer.|

## Remarks

Use the **Hourglass** method to change the mouse pointer to an image of an hourglass (or another icon you've chosen) while a procedure is running. This method can provide a visual indication that the procedure is running. This is especially useful when a procedure takes a long time to run.

You often use this method if you've turned echo off by using the **Echo** method. When echo is off, Microsoft Access suspends screen updates until the macro is finished.

Access automatically resets the _HourglassOn_ argument to **False** when the procedure finishes running.

To determine the current state of the hourglass, you can check the value of [Screen.MousePointer](/office/vba/api/access.screen.mousepointer). If `Screen.MousePointer = 11`, the hourglass is currently being displayed.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
