---
title: UserProperties.Find method (Outlook)
keywords: vbaol11.chm210
f1_keywords:
- vbaol11.chm210
ms.prod: outlook
api_name:
- Outlook.UserProperties.Find
ms.assetid: 3b71ce5a-4bb0-fdab-a24e-02c631816b80
ms.date: 06/08/2017
localization_priority: Normal
---


# UserProperties.Find method (Outlook)

Locates and returns a **[UserProperty](Outlook.UserProperty.md)** object for the requested property name, if it exists.


## Syntax

_expression_.**Find** (_Name_, _Custom_)

_expression_ A variable that represents a **[UserProperties](Outlook.UserProperties.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the requested property.|
| _Custom_|Optional| **Variant**| **True** if custom properties on the item should be searched, **False** if built-in properties should be searched.|

## Return value

If you use **Find** to look for a custom property and the call succeeds, it returns a **UserProperty** object. If it fails, it returns **Null** (**Nothing** in Visual Basic). 

If you use **Find** to look for a built-in property, specify **False** for the _Custom_ parameter. If the call succeeds, it returns the property as a **UserProperty** object. If the call fails, it returns **Null** (**Nothing** in Visual Basic). If you specify **True** for _Custom_, the call does not find the built-in property and returns **Null** (**Nothing** in Visual Basic).


## Remarks

If the _Custom_ parameter is **True**, only custom user properties are searched. The default value is **True**. To find a non-custom property such as **Subject**, specify the _Custom_ parameter as **False**; otherwise, it returns **Nothing**.


## See also


[UserProperties Object](Outlook.UserProperties.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]