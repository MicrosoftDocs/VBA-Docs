---
title: Replica has not been synchronized within the replica set retention period. (Error 3743)
keywords: jeterr40.chm5003743
f1_keywords:
- jeterr40.chm5003743
ms.assetid: 52fd5406-2664-8cbe-f1ac-f37c3cb7ad5c
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Replica has not been synchronized within the replica set retention period. (Error 3743)

  

**Applies to:** Access 2013 | Access 2016

If the retention period expires for a replica, you cannot synchronize changes between the expired replica and the other replicas in the replica set. If a replica does not synchronize with another replica in the set within the retention period, the next time you attempt to synchronize the replica it gets removed from the replica set. The retention period is established when the database is initially made replicable. If you replicate the database by using Replication Manager, Data Access Objects (DAO), or ActiveX Data Objects (ADO), the default retention period is 60 days. If you replicate the database by using Microsoft Access or Briefcase, the default retention period is 1000 days.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]