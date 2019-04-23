---
title: Insert, update, and delete records from a table using Access SQL
ms.prod: access
ms.assetid: 0d71f4f1-efc1-127e-5edc-263a3a2a30fb
ms.date: 06/08/2017
localization_priority: Priority
---


# Insert, update, and delete records from a table using Access SQL

## Insert records into a table

There are essentially two methods for adding records to a table. The first is to add one record at a time; the second is to add many records at a time. In both cases, you use the SQL statement **[INSERT INTO](../../../api/overview/Access.md)** to accomplish the task. **INSERT INTO** statements are commonly referred to as append queries.

To add one record to a table, you must use the field list to define which fields to put the data in, and then you must supply the data itself in a value list. To define the value list, use the **VALUES** clause. For example, the following statement will insert the values "1", "Kelly", and "Jill" into the CustomerID, Last Name, and First Name fields, respectively.


```sql
INSERT INTO tblCustomers (CustomerID, [Last Name], [First Name]) 
    VALUES (1, 'Kelly', 'Jill') 

```

You can omit the field list, but only if you supply all the values that record can contain.

```sql
INSERT INTO tblCustomers 
    VALUES (1, Kelly, 'Jill', '555-1040', 'someone@microsoft.com') 

```

To add many records to a table at one time, use the **INSERT INTO** statement along with a **[SELECT](../../../api/overview/Access.md)** statement. When you are inserting records from another table, each value being inserted must be compatible with the type of field that will be receiving the data.

The following **INSERT INTO** statement inserts all the values in the CustomerID, Last Name, and First Name fields from the tblOldCustomers table into the corresponding fields in the tblCustomers table.


```sql
INSERT INTO tblCustomers (CustomerID, [Last Name], [First Name]) 
    SELECT CustomerID, [Last Name], [First Name] 
    FROM tblOldCustomers 

```

If the tables are defined exactly alike, you can leave out the field lists.

```sql
INSERT INTO tblCustomers 
    SELECT * FROM tblOldCustomers 

```


## Update records in a table

To modify the data that is currently in a table, you use the **[UPDATE](../../../api/overview/Access.md)** statement, which is commonly referred to as an update query. The **UPDATE** statement can modify one or more records and generally takes this form.


```sql
UPDATE table name   
    SET field name  = some value
```

To update all the records in a table, specify the table name, and then use the **SET** clause to specify the field or fields to be changed.




```sql
UPDATE tblCustomers 
    SET Phone = 'None' 

```

In most cases, you will want to qualify the **UPDATE** statement with a **[WHERE](where-clause-microsoft-access-sql.md)** clause to limit the number of records changed.




```sql
UPDATE tblCustomers 
    SET Email = 'None' 
    WHERE [Last Name] = 'Smith' 

```


## Delete records from a table

To delete the data that is currently in a table, you use the **[DELETE](../../../api/overview/Access.md)** statement, which is commonly referred to as a delete query. This is also known as truncating a table. The **DELETE** statement can remove one or more records from a table and generally takes this form:


```sql
DELETE FROM table list
```

The **DELETE** statement does not remove the table structure—only the data that is currently being held by the table structure. To remove all the records from a table, use the **DELETE** statement and specify which table or tables from which you want to delete all the records.




```sql
DELETE FROM tblInvoices 

```

In most cases, you will want to qualify the **DELETE** statement with a **WHERE** clause to limit the number of records to be removed.




```sql
DELETE FROM tblInvoices 
    WHERE InvoiceID = 3 

```

If you want to remove data only from certain fields in a table, use the **UPDATE** statement and set those fields equal to **NULL**, but only if they are nullable fields.




```sql
UPDATE tblCustomers  
    SET Email = Null 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
