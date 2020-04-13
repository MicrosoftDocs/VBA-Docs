---
title: Illegal column-level constraint. (Error 3789)
keywords: jeterr40.chm5003789
f1_keywords:
- jeterr40.chm5003789
ms.prod: access
ms.assetid: 66b78a40-bfc6-28dd-77b3-0c876a163f25
ms.date: 06/08/2019
localization_priority: Normal
---


# Illegal column-level constraint. (Error 3789)

  

**Applies to:** Access 2013 | Access 2016

This error occurs when using the CREATE TABLE or ALTER TABLE ALTER COLUMN syntax. While ANSI SQL allows for creating CHECK constraints as part of the table definition, the Microsoft Access database engine requires that the user create the CHECK constraint separate from the COLUMN definition. This can be accomplished by issuing the CHECK keyword after a comma. For example, the following syntax will work because the CHECK constraint is defined separately from the column and follows a comma:

CREATE TABLE Orders (OrderId IDENTITY (100,10) CONSTRAINT pkOrders PRIMARY KEY, CustId LONG CONSTRAINT fkCustomersCustId REFERENCES Customers (CustId), Balance DOUBLE, CONSTRAINT CustomerExceededCreditLimit CHECK (CustId IN (SELECT CustId FROM Customers C WHERE C.CustId = Orders.CustId AND C.CreditLimit >= (SELECT SUM(Balance)FROM Orders O WHERE O.CustId = Orders.CustId))));.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]