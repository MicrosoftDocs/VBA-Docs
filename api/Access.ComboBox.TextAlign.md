---
title: ComboBox.TextAlign property (Access)
keywords: vbaac10.chm11420
f1_keywords:
- vbaac10.chm11420
ms.prod: access
api_name:
- Access.ComboBox.TextAlign
ms.assetid: c5de59ad-f41f-8f19-6056-16ca88a1937d
ms.date: 03/02/2019
localization_priority: Normal
---


# ComboBox.TextAlign property (Access)

The **TextAlign** property specifies the text alignment in new controls. Read/write **Byte**.


## Syntax

_expression_.**TextAlign**

_expression_ A variable that represents a **[ComboBox](Access.ComboBox.md)** object.


## Remarks

The **TextAlign** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|General|0|(Default) The text aligns to the left; numbers and dates align to the right.|
|Left|1|The text, numbers, and dates align to the left.|
|Center|2|The text, numbers, and dates are centered.|
|Right|3|The text, numbers, and dates align to the right.|
|Distribute|4|The text, numbers, and dates are evenly distributed.|

You can set the default for the **TextAlign** property by using a control's default control style or the **[DefaultControl](access.form.defaultcontrol.md)** property in Visual Basic.


## Example

The following example aligns the text in the **Address** text box on the **Suppliers** form to the right.

```vb
Forms("Suppliers").Controls("Address").TextAlign = 3
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]