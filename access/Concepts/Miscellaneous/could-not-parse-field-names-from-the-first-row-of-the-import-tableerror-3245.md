---
title: Could not parse field names from the first row of the import table. (Error 3245)
keywords: jeterr40.chm5003245
f1_keywords:
- jeterr40.chm5003245
ms.assetid: ac70f60f-e43b-30cc-fea4-969c132819df
ms.date: 02/14/2020
ms.localizationpriority: medium
---

# Could not parse field names from the first row of the import table. (Error 3245)

**Applies to:** Access 2013 | Access 2016

The first row of data contains invalid field names, such as quoted and unquoted strings in the same field name. In the following example, the third and fourth field names cannot be parsed:

`"Name", Date, "ID " Number, Phone" Number"`

The first two fields are valid, but the third and fourth are not because they contain nonspace characters outside the quotation marks.
Check the import table for properly matched quotation marks, and then try the import operation again

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
