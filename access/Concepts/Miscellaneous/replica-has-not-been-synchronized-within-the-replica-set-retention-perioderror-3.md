---
title: Replica has not been synchronized within the replica set retention period. (Error 3743)
keywords: jeterr40.chm5003743
f1_keywords:
- jeterr40.chm5003743
ms.prod: access
ms.assetid: 52fd5406-2664-8cbe-f1ac-f37c3cb7ad5c
ms.date: 06/08/2017
localization_priority: Normal
---


# Replica has not been synchronized within the replica set retention period. (Error 3743)

  

**Applies to:** Access 2013 | Access 2016

If the retention period expires for a replica, you cannot synchronize changes between the expired replica and the other replicas in the replica set. If a replica does not synchronize with another replica in the set within the retention period, the next time you attempt to synchronize the replica it gets removed from the replica set. The retention period is established when the database is initially made replicable. If you replicate the database by using Replication Manager, Data Access Objects (DAO), or ActiveX Data Objects (ADO), the default retention period is 60 days. If you replicate the database by using Microsoft Access or Briefcase, the default retention period is 1000 days.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]