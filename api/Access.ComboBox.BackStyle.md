---
title: ComboBox.BackStyle property (Access)
keywords: vbaac10.chm11405
f1_keywords:
- vbaac10.chm11405
ms.prod: access
api_name:
- Access.ComboBox.BackStyle
ms.assetid: 1def822f-6b4a-8384-9d81-72b30e680908
ms.date: 02/28/2019
localization_priority: Normal
---


# ComboBox.BackStyle property (Access)

You can use the **BackStyle** property to specify whether a control will be transparent. Read/write **Byte**.


## Syntax

_expression_.**BackStyle**

_expression_ A variable that represents a **[ComboBox](Access.ComboBox.md)** object.


## Remarks

The **BackStyle** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|Normal|1|(Default for all controls except option group) The control has its interior color set by the **BackColor** property.|
|Transparent|0|(Default for option group) The control is transparent. The color of the form or report behind the control is visible.|

You can set the default for this property by using a control's default control style or the **[DefaultControl](access.form.defaultcontrol.md)** property in Visual Basic.

If the **Transparent** button on the **Back Color** button palette is selected, the **BackStyle** property is set to Transparent; otherwise, the **BackStyle** property is set to Normal.

To make a command button invisible, set its **Transparent** property to Yes.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]