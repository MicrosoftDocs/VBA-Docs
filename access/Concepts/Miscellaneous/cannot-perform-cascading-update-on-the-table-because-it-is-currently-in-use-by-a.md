---
title: Cannot perform cascading update on the table because it is currently in use by another user. (Error 3412)
keywords: jeterr40.chm5003412
f1_keywords:
- jeterr40.chm5003412
ms.prod: access
ms.assetid: 0718b58e-5553-8c08-ea85-83f97eb88998
ms.date: 06/08/2017
---


# Cannot perform cascading update on the table because it is currently in use by another user. (Error 3412)

  

**Applies to:** Access 2013 | Access 2016

You are trying to save a value into a primary key field that is included in a relationship.

In Microsoft Access, the  **Cascade Update Related Fields** option is selected for the relationship, or in DAO code, the **dbRelationUpdateCascade** option is specified for the **Relation** object's **Attributes** property. Your application is attempting to update the matching field in the related table.
The matching field cannot be updated because of a locking conflict with another user. To save the record, you must wait until the related table is no longer in use.
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](https://answers.microsoft.com/)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

