---
title: Ignore Nulls property
keywords: acmain11.chm7025
f1_keywords:
- acmain11.chm7025
ms.prod: access
ms.assetid: 87d95ca8-ea29-f0ca-366a-56527c500f13
ms.date: 06/08/2017
localization_priority: Normal
---


# Ignore Nulls property

**Applies to:** Access 2013 | Access 2016

You can use the IgnoreNulls property to specify that records with Null values in the indexed fields not be included in the index.

## Settings

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|Yes|**True**|Records that contain Null values in the indexed fields aren't included in the index.|
|No|**False**|(Default) Records that contain Null values in the indexed fields are included in the index.|
You can set this property by using the Indexes window of table Design view or Visual Basic.

To access the  **Ignore Nulls** property of an index by using Visual Basic, use the DAO **IgnoreNulls** property.

You can define an index for a field to facilitate faster searches for records indexed on that field. If you allow  **Null** entries in the indexed field and expect to have many of them, set the **Ignore Nulls** property for the index to Yes to reduce the amount of storage space that the index uses.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]