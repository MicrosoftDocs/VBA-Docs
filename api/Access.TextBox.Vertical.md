---
title: TextBox.Vertical property (Access)
keywords: vbaac10.chm11058
f1_keywords:
- vbaac10.chm11058
ms.prod: access
api_name:
- Access.TextBox.Vertical
ms.assetid: 40b9f9c0-daab-5562-395e-3e785d316d91
ms.date: 03/26/2019
localization_priority: Normal
---


# TextBox.Vertical property (Access)

You can use the **Vertical** property to set a form control for vertical display and editing, or to set a report control for vertical display and printing. Read/write **Boolean**.


## Syntax

_expression_.**Vertical**

_expression_ A variable that represents a **[TextBox](Access.TextBox.md)** object.


## Remarks

The **Vertical** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|Yes|**True**|Displays, edits, and prints vertical text.|
|No|**False**|(Default) Does not display, edit, or print vertical text.|

You can specify how vertical text will be displayed, edited, or printed in the control by setting the **Vertical** property. If set to Yes, the starting point for inputting text is the upper-right corner of the control (the ending point is the lower-left corner of the control). 

If using full pitch characters, the display and print directions are the same as the control for horizontal text. If using half pitch characters, it shifts 90 degrees to the right. The cursor is also rotated 90 degrees to the right in a vertical text control.

Text selection using key combinations is different for vertical text control and horizontal text control. Key combinations and their effect on range selection are described in the following table.

|Key combination|Vertical text control|Horizontal text control|
|:--------------|:--------------------|:----------------------|
|Shift+Up|One character before the cursor. |One line before the cursor.|
|Shift+Down|One character after the cursor. |One line after the cursor.|
|Shift+Right|One line after the cursor. |One character before the cursor.|
|Shift+Left|One line before the cursor. |One character after the cursor.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]