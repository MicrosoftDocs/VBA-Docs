---
title: ComboBox.SelectionMargin Property (Outlook Forms Script)
keywords: olfm10.chm2001860
f1_keywords:
- olfm10.chm2001860
ms.prod: outlook
ms.assetid: 68d1b2c3-950b-1928-a790-edfbbc5de4b5
ms.date: 06/08/2017
localization_priority: Normal
---


# ComboBox.SelectionMargin Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether the user can select a line of text by clicking in the region to the left of the text. Read/write.


## Syntax

_expression_.**SelectionMargin**

_expression_ A variable that represents a  **ComboBox** object.


## Remarks

 **True** if clicking in margin causes selection of text (default), **False** if clicking in margin does not cause selection of text.

When the  **SelectionMargin** property is **True**, the selection margin occupies a thin strip along the left edge of a control's edit region. When set to  **False**, the entire edit region can store text.

If the  **SelectionMargin** property is set to **True** when a control is printed, the selection margin also prints.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]