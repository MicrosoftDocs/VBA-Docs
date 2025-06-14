---
title: Attachment.OldBorderStyle property (Access)
keywords: vbaac10.chm13928
f1_keywords:
- vbaac10.chm13928
api_name:
- Access.Attachment.OldBorderStyle
ms.assetid: abbc1a8d-d9cc-b917-026d-a1847739c362
ms.date: 02/07/2019
ms.localizationpriority: medium
---


# Attachment.OldBorderStyle property (Access)

Use this property to set or return the unedited value of the **[BorderStyle](access.attachment.borderstyle.md)** property for a form or control. This property is useful if you need to revert to an unedited or preferred border style. Read/write **Byte**.


## Syntax

_expression_.**OldBorderStyle**

_expression_ A variable that represents an **[Attachment](Access.Attachment.md)** object.


## Remarks

The **OldBorderStyle** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|Transparent|0|Transparent|
|Solid|1|(Default) Solid line|
|Dashes|2|Dashed line|
|Short dashes|3|Dashed line with short dashes|
|Dots|4|Dotted line|
|Sparse dots|5|Dotted line with dots spaced far apart|
|Dash dot|6|Line with a dash-dot combination|
|Dash dot dot|7|Line with a dash-dot-dot combination|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]