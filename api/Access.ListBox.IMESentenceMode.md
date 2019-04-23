---
title: ListBox.IMESentenceMode property (Access)
keywords: vbaac10.chm11298
f1_keywords:
- vbaac10.chm11298
ms.prod: access
api_name:
- Access.ListBox.IMESentenceMode
ms.assetid: 877e1766-c378-cf7b-b452-bb8f536980f3
ms.date: 03/01/2019
localization_priority: Normal
---


# ListBox.IMESentenceMode property (Access)


## Syntax

_expression_.**IMESentenceMode**

_expression_ A variable that represents a **[ListBox](Access.ListBox.md)** object.


## Remarks

The **IMESentenceMode** property uses the following settings.

|Setting|Description|Visual Basic|
|:-----|:-----|:-----|
|Normal|(Default) Sets IME Sentence Mode to Normal mode.<br/><br/>Use this mode when creating a literary Japanese document.|0|
|Plural|Sets IME Sentence Mode to Plural mode.<br/><br/>Use this mode when entering name or address data. In this mode, two additional dictionaries are available. The Biographical/Geographical Dictionary contains names not covered in the normal dictionary, and the Postal Code Dictionary is useful in creating addresses (factory setting).|1|
|Speaking|Sets IME Sentence Mode to Speaking mode.<br/><br/>Use this mode when entering data that contains conversational language.|2|
|No Conversion|Doesn't set IME Sentence Mode.<br/><br/>In this mode, inputted characters are settled without conversion.|3|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]