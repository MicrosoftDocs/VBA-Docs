---
title: AllowZeroLength property
ROBOTS: INDEX
keywords: vbaac10.chm4268
f1_keywords:
- vbaac10.chm4268
ms.prod: access
api_name:
- Access.AllowZeroLength
ms.assetid: e65dd834-7daa-ad78-0560-87aad5efa1a8
ms.date: 06/08/2019
localization_priority: Normal
---


# AllowZeroLength property

**Applies to:** Access 2013 | Access 2016

You can use the **AllowZeroLength** property to specify whether a zero-length string (" ") is a valid entry in a table field.

> [!NOTE] 
> The **AllowZeroLength** property applies only to Text, Memo, and Hyperlink table fields.


## Setting

The **AllowZeroLength** property uses the following settings.


|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|Yes|**True**|A zero-length string is a valid entry. This is the default value when creating a field in the Access user interface.|
|No|**False**|A zero-length string is an invalid entry. This is the default value when creating a field programmatically.|

> [!NOTE] 
> To access a field's **AllowZeroLength** property by using Visual Basic, use the DAO **AllowZeroLength** property or the ADO **Column.Properties** ("Set OLEDB:Allow Zero Length") property.


## Remarks

If you want Microsoft Access to store a zero-length string instead of a Null value when you leave a field blank, set both the **AllowZeroLength** and **Required** properties to **Yes**.

The following table shows the results of combining the settings of the **AllowZeroLength** and **Required** properties.

|**AllowZeroLength**|**Required**|**User's Action**|**Value Stored**|
|:-----|:-----|:-----|:-----|
|No|No|Presses ENTER<br/>Presses SPACEBAR<br/>Enters a zero-length string|**Null**<br/>**Null**<br/>(not allowed)|
|Yes|No|Presses ENTER<br/>Presses SPACEBAR<br/>Enters a zero-length string|**Null**<br/>**Null**<br/>Zero-length string|
|No|Yes|Presses ENTER<br/>Presses SPACEBAR<br/>Enters a zero-length string|(not allowed)<br/>(not allowed)<br/>(not allowed)|
|Yes|Yes|Presses ENTER<br/>Presses SPACEBAR<br/>Enters a zero-length string|(not allowed)<br/>Zero-length string<br/>Zero-length string|

> [!NOTE] 
> You can use the **Format** property to distinguish between the display of a **Null** value and a zero-length string. For example, the string "None" can be displayed when a zero-length string is entered.

The **AllowZeroLength** property works independently of the **Required** property. The **Required** property determines only whether a **Null** value is valid for the field. If the **AllowZeroLength** property is set to Yes, a zero-length string will be a valid value for the field regardless of the setting of the **Required** property.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]