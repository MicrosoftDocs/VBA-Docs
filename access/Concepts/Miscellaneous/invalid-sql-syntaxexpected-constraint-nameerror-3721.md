---
title: Invalid SQL syntax - expected CONSTRAINT name. (Error 3721)
ms.assetid: 14da04b2-b7d0-3e23-20fe-20e42ef4b3d7
ms.date: 02/14/2019
ms.localizationpriority: medium
---


# Invalid SQL syntax - expected CONSTRAINT name. (Error 3721)

**Applies to:** Access 2013 | Access 2016

When defining referential integrity from a SQL DDL statement, it is necessary to name a constraint when using the CONSTRAINT keyword. If a constraint name is not desired, don't use the CONSTRAINT keyword. An example of this error would be:

```sql
CREATE TABLE Customers (CLstNm TEXT(50), CFrstNm TEXT(25), CONSTRAINT PRIMARY KEY (CFrstNm, CLstNm));
```

To prevent the error, include a name after the CONSTRAINT keyword:

```sql
CREATE TABLE Customers (CLstNm TEXT(50), CFrstNm TEXT(25), CONSTRAINT pkCustomers PRIMARY KEY (CFrstNm, CLstNm));
```

or don't use the CONSTRAINT keyword:

```sql
CREATE TABLE Customers (CLstNm TEXT(50), CFrstNm TEXT(25), PRIMARY KEY (CFrstNm, CLstNm));
```

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]