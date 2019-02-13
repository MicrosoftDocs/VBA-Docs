---
title: Invalid SQL syntax - expected CONSTRAINT name. (Error 3721)
ms.prod: access
ms.assetid: 14da04b2-b7d0-3e23-20fe-20e42ef4b3d7
ms.date: 02/14/2019
localization_priority: Normal
---


# Invalid SQL syntax - expected CONSTRAINT name. (Error 3721)

**Applies to:** Access 2013 | Access 2016

When defining referential integrity from a SQL DDL statement, it is necessary to name a constraint when using the CONSTRAINT keyword. If a constraint name is not desired, do not use the CONSTRAINT keyword. An example of this error would be:

```sql
CREATE TABLE Customers (CLstNm TEXT(50), CFrstNm TEXT(25), CONSTRAINT PRIMARY KEY (CFrstNm, CLstNm));
```

<br/>

To prevent the error, include a name after the CONSTRAINT keyword:

```sql
CREATE TABLE Customers (CLstNm TEXT(50), CFrstNm TEXT(25), CONSTRAINT pkCustomers PRIMARY KEY (CFrstNm, CLstNm));
```

or do not use the CONSTRAINT keyword:

```sql
CREATE TABLE Customers (CLstNm TEXT(50), CFrstNm TEXT(25), PRIMARY KEY (CFrstNm, CLstNm));
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]