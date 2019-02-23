---
title: TextBox.ReadingOrder property (Access)
keywords: vbaac10.chm11130
f1_keywords:
- vbaac10.chm11130
ms.prod: access
api_name:
- Access.TextBox.ReadingOrder
ms.assetid: 1b53bb00-9252-ca99-c3b7-3a97d06552c4
ms.date: 02/23/2019
localization_priority: Normal
---


# TextBox.ReadingOrder property (Access)

You can use the **ReadingOrder** property to specify or determine the reading order of words in text. Read/write **Byte**.


## Syntax

_expression_.**ReadingOrder**

_expression_ A variable that represents a **[TextBox](Access.TextBox.md)** object.


## Remarks

The **ReadingOrder** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|Context|0|Reading order is determined by the language of the first character entered.<br/><br/>If a right-to-left language character is entered first, reading order is right to left.<br/><br/>If a left-to-right language character is entered first, reading order is left to right.|
|Left-to-Right|1|Sets the reading order to left to right.|
|Right-to-Left|2|Sets the reading order to right to left.|


## Example

The following example sets the reading order to right to left for the **Address** text box on the **International Shipping** form.

```vb
Forms("International Shipping").Controls("Address").ReadingOrder = 2
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

