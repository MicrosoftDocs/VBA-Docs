---
title: Synchronizing with a non-replicated database is not allowed. The <name> database is not a Design Master or replica. (Error 3605)
keywords: jeterr40.chm5003605
f1_keywords:
- jeterr40.chm5003605
ms.prod: access
ms.assetid: 5233d276-9a31-bbe3-7b2e-33636c7df521
ms.date: 06/08/2017
localization_priority: Normal
---


# Synchronizing with a non-replicated database is not allowed. The <name> database is not a Design Master or replica. (Error 3605)

  

**Applies to:** Access 2013 | Access 2016

You are attempting to synchronize a replicated database with a non-replicated database or to synchronize two non-replicated databases. Only replicas made from the same replicated database can be synchronized.

If one of the databases has already been replicated, use it to create your second database replica.
If neither database has been replicated, select one of the databases to be used as the Design Master for the replica set. Open that database using Microsoft Access, go to the  **Tools** menu, point to **Replication**, and click **Create Replica**. If Microsoft Access is not available but Microsoft Windows 95 Briefcase is available, drag the database into the Briefcase to create a replica. Do not attempt to replicate the second of the two original databases and then synchronize the two databases. The second database must be a replica of the first for synchronization to succeed.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]