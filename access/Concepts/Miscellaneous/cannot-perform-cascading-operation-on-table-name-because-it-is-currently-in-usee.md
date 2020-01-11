---
title: Cannot perform cascading operation on table <name> because it is currently in use. (Error 3414)
keywords: jeterr40.chm5003414
f1_keywords:
- jeterr40.chm5003414
ms.prod: access
ms.assetid: 238227df-7dd2-6a72-7c3d-8b76a5bc7834
ms.date: 06/08/2017
localization_priority: Normal
---


# Cannot perform cascading operation on table <name> because it is currently in use. (Error 3414)

  

**Applies to:** Access 2013 | Access 2016

You are trying to save a value to a primary key field that is included in a relationship.

In Microsoft Access, the  **Cascade Update Related Fields** option is selected for the relationship, or in DAO code, the **dbRelationUpdateCascade** option is specified for the **Relation** object's **Attributes** property. Therefore, your application is attempting to update the matching field in the related table.
The matching field cannot be updated because you have it open or locked on your computer. To save the record, you must first close the related table.


## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]