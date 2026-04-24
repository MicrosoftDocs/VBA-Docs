---
title: Cannot create relationships to enforce referential integrity. Existing data in table <name> violates referential integrity rules in table <name>. (Error 3379)
keywords: jeterr40.chm5003379
f1_keywords:
- jeterr40.chm5003379
ms.assetid: 2206ce0e-447f-edda-dadf-c931d3e5f834
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Cannot create relationships to enforce referential integrity. Existing data in table \<name\> violates referential integrity rules in table \<name\>. (Error 3379)

  

**Applies to:** Access 2013 | Access 2016

You are trying to create a relationship using the CONSTRAINT clause of the ALTER TABLE statement, but existing data in the two tables violates referential integrity constraints. For example, there might be records relating to an employee in the related table but no corresponding record for the employee in the primary table.

To create the relationship, you must edit the data so that primary records exist for all related records.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]