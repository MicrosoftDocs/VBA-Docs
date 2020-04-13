---
title: This member of the replica set has been logically removed from the set and is no longer available. (Error 3569)
keywords: jeterr40.chm5003569
f1_keywords:
- jeterr40.chm5003569
ms.prod: access
ms.assetid: 5457a862-fee7-14d1-b9a9-0967cdabec28
ms.date: 06/08/2017
localization_priority: Normal
---


# This member of the replica set has been logically removed from the set and is no longer available. (Error 3569)

  

**Applies to:** Access 2013 | Access 2016

The member of the replica set you are attempting to use has been removed because either it has not been synchronized within the required number of days (retention period) or it has been synchronized with a second Design Master. The member should be deleted from your computer.

If there is data in the member that has not been synchronized with any other member in the replica set, manually enter that data in another member. You can determine which records have unsynchronized data by examining the s_Generation field in each table. A value of 0 in the s_Generation field indicates the record is an inserted or updated and has not been synchronized with another replica in the set. To view the s_Generation field, open Microsoft Access, go to the **Tools** menu, click **Options**, select the **View** tab, and then select the **System Objects** check box.
If the member is the Design Master, compare the structure of the replica with another replica in the set to determine whether any design changes have not been synchronized.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]