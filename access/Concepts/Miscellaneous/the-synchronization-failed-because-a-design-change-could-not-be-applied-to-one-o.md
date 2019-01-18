---
title: The synchronization failed because a design change could not be applied to one of the replicas. (Error 3492)
keywords: jeterr40.chm5003492
f1_keywords:
- jeterr40.chm5003492
ms.prod: access
ms.assetid: 08ba127a-7002-84ae-6f76-65f4aedeb052
ms.date: 06/08/2017
localization_priority: Normal
---


# The synchronization failed because a design change could not be applied to one of the replicas. (Error 3492)

  

**Applies to:** Access 2013 | Access 2016

The Microsoft Access database engine attempted to update the database design at one of the replicas. There are several possible reasons why the design could not be updated, including:



- The object you are trying to update is already open at the replica.
    
- You added an enforced relationship to a replica that has a foreign key that references a nonexistent primary key.
    

For additional information regarding the synchronization failure, look in the MSysSchemaProb table, either at the Design Master or the replica that was the target of the synchronization.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]