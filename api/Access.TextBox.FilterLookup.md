---
title: TextBox.FilterLookup property (Access)
keywords: vbaac10.chm11062,vbaac10.chm4353
f1_keywords:
- vbaac10.chm11062,vbaac10.chm4353
ms.prod: access
api_name:
- Access.TextBox.FilterLookup
ms.assetid: 5c568366-94a5-8d7a-1fb4-80b4b3ab6c7f
ms.date: 03/26/2019
localization_priority: Normal
---


# TextBox.FilterLookup property (Access)

You can use the **FilterLookup** property to specify whether values appear in a bound text box control when using the Filter By Form or Server Filter By Form window. Read/write **Byte**.


## Syntax

_expression_.**FilterLookup**

_expression_ A variable that represents a **[TextBox](Access.TextBox.md)** object.


## Remarks

The **FilterLookup** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|Never|0|The field values aren't displayed. You can specify whether the filtered records can contain null values.|
|Database Default|1|(Default) The field values are displayed according to the settings under **Filter lookup options** on the **Current Database** tab of the **Access Options** dialog box, available by choosing the Microsoft Office button, and then choosing **Access Options**.|
|Always|2|The field values are always displayed.|

You can also set the default for this property by setting a control's **[DefaultControl](access.form.defaultcontrol.md)** property in Visual Basic.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]