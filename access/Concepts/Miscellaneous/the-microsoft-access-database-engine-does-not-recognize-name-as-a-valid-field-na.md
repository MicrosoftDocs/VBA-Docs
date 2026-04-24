---
title: The Microsoft Access database engine does not recognize <name> as a valid field name or expression. (Error 3070)
keywords: jeterr40.chm5003070
f1_keywords:
- jeterr40.chm5003070
ms.assetid: 8866f9ea-4c2b-45f6-9ec7-8e23596efbf9
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# The Microsoft Access database engine does not recognize \<name\> as a valid field name or expression. (Error 3070)

**Applies to:** Access 2013 | Access 2016

The specified name is not a recognized field name or a valid expression. In a query, this error can occur if you enter a name that improperly refers to a database, table, or field.

Possible causes with Microsoft Access:

- You have a parameter in a crosstab query or in a query that a crosstab query or chart is based on, and the parameter data type is not explicitly specified in the **Query Parameters** dialog box. To solve the problem:
    
  - In the query that contains the parameter, specify the parameter and its data type in the **Query Parameters** dialog box. And;
    
  - Set the **ColumnHeadings** property for the query that contains the parameter.  
        
  - In any type of query, you have improperly referred to a database, table, or field. For example, this error can occur if you refer to a field named Salary in an expression, but you misspell the field name, such as `[Sallary]*1.1`.
    

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
