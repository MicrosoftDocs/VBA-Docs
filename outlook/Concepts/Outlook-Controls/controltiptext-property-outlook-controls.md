---
title: ControlTipText Property (Outlook Controls)
ms.prod: outlook
ms.assetid: 8dac3e44-f25c-b1b9-8347-86fd7e688e81
ms.date: 06/08/2017
localization_priority: Normal
---


# ControlTipText Property (Outlook Controls)

Returns or sets a  **String** that appears when the user briefly holds the mouse pointer over a control without clicking. Read/write.


## Syntax

 _expression_. **ControlTipText**

 _expression_ A variable that represents an Outlook control object.


## Remarks

The  **ControlTipText** property lets you give users tips about a control in a running form. The property can be set during design time but only appears by the control during runtime.

The default value of  **ControlTipText** is an empty string. When the value of **ControlTipText** is set to an empty string, no tip is available for that control.

Note that for the  **[OlkBusinessCardControl](../../../api/Outlook.OlkBusinessCardControl.md)** and **[OlkContactPhoto](../../../api/Outlook.OlkContactPhoto.md)** controls, **ControlTipText** is not displayed when mousing over the control.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]