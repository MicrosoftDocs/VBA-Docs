---
title: After a database has been replicated, you cannot remove the replication features from the database. (Error 3459)
keywords: jeterr40.chm5003459
f1_keywords:
- jeterr40.chm5003459
ms.assetid: 06429f0d-7b78-d7e7-3c67-23142282b0ca
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# After a database has been replicated, you cannot remove the replication features from the database. (Error 3459)

  

**Applies to:** Access 2013 | Access 2016

You cannot remove the replication features from a database that has been replicated, either by setting its **Replicable** property to "T" or by converting it with the Microsoft Windows Briefcase, Microsoft Access, or the Replication Manager. Using DAO to set the database's **Replicable** property to "F" or any other value has no effect on the replicated database. If you need to use a copy of the database that does not have the properties, fields, tables, and other characteristics associated with replication, open the backup copy of the database made before the database was first converted.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]