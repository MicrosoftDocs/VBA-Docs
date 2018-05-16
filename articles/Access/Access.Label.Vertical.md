---
title: Label.Vertical Property (Access)
keywords: vbaac10.chm10196
f1_keywords:
- vbaac10.chm10196
ms.prod: access
api_name:
- Access.Label.Vertical
ms.assetid: 6ce97069-0713-9a6f-3efc-4a5161ee54e3
ms.date: 06/08/2017
---


# Label.Vertical Property (Access)

You can use the  **Vertical** property to set a form control for vertical display and editing or set a report control for vertical display and printing. Read/write **Boolean**.


## Syntax

 _expression_. **Vertical**

 _expression_ A variable that represents a **Label** object.


## Remarks

The  **Vertical** property uses the following settings:



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|Displays, edits, and prints vertical text.|
|No|**False**|(Default) Does not display, edit, or print vertical text. |
You can specify how vertical text will be displayed, edited, or printed in the control by setting the  **Vertical** property. If set to Yes, the starting point for inputting text is the upper right corner of the control (the ending point is the lower left corner of the control). If using full pitch characters, the display and print directions are the same as the control for horizontal text. If using half pitch characters, it shifts 90 degrees to the right. The cursor is also rotated 90 degrees to the right in a vertical text control.


 **Note**  Text selection using key combinations is different for vertical text control and horizontal text control. Key combinations and their effect on range selection are described below.



|**Key combination**|**Selected range**|
|:-----|:-----|
|Shift+Up|Vertical: One character before the cursor. Horizontal: One line before the cursor.|
|Shift+Down|Vertical: One character after the cursor. Horizontal: One line after the cursor.|
|Shift+Right|Vertical: One line after the cursor. Horizontal: One character before the cursor.|
|Shift+Left|Vertical: One line before the cursor. Horizontal: One character after the cursor.|

## See also


#### Concepts


[Label Object](Access.Label.md)

