---
title: Handler property example (VJ++)
ROBOTS: INDEX
ms.prod: access
ms.assetid: fba66f04-654d-5950-ee48-0da6f6106be2
ms.date: 06/08/2019
localization_priority: Normal
---


# Handler property example (VJ++)

**Applies to:** Access 2013 | Access 2016

This example demonstrates the [RDS DataControl](https://msdn.microsoft.com/library/ac430669-7628-696c-c036-b5d35405d788%28Office.15%29.aspx) object [Handler](https://msdn.microsoft.com/library/aaf8c8c6-f95b-3cf3-b3f6-203f37464c87%28Office.15%29.aspx) property. (See [DataFactory Customization](https://msdn.microsoft.com/library/43cd7416-1f05-87ee-22f0-6cf0d2d1b39f%28Office.15%29.aspx) for more details.)

Assume that the following sections in the parameter file, Msdfmap.ini, are located on the server:

```ini
 
[connect AuthorDataBase] 
Access=ReadWrite 
Connect="DSN=Pubs" 
[sql AuthorById] 
SQL="SELECT * FROM Authors WHERE au_id = ?" 

```

Your code looks like the following. The command assigned to the [SQL](sql-property-ado.md) property will match the _AuthorById_ identifier and will retrieve a row for author Michael O'Leary. Although the [Connect](https://msdn.microsoft.com/library/11aa3284-18e9-6d2d-761b-c25090370b77%28Office.15%29.aspx) property in your code specifies the Northwind data source, that data source will be overwritten by the Msdfmap.ini _connect_ section. The **DataControl** object's [Recordset](https://msdn.microsoft.com/library/5f4bb72d-ddfa-41c0-c353-b3a6632b4a91%28Office.15%29.aspx) property is assigned to a disconnected [Recordset](https://msdn.microsoft.com/library/0f963bf8-f066-dc8a-b754-f427de712df1%28Office.15%29.aspx) object purely as a coding convenience.


```csharp

// BeginHandlerJimport com.ms.wfc.data.*;
import com.ms.wfc.data.rds.*;import java.io.* ;
public class HandlerX{
// The main entry point for the application.public static void main (String[] args)
{HandlerX();
System.exit(0);}
// HandlerX functionstatic void HandlerX()
{// Define ADO Objects.
Recordset rstAuthors = null;// Declarations.
BufferedReader in =new BufferedReader (new InputStreamReader(System.in));
int intCount = 0;int intDisplaysize = 15;
try{
IBindMgr dc = (IBindMgr) new DataControl();dc.setServer("MyServer");
dc.setConnect("Data Source=Northwind");dc.setSQL("AuthorById(267-41-2394)");
dc.Refresh(); // Retrieve the record.// Use another recordset as a convenience.
rstAuthors = (Recordset)dc.getRecordset();System.out.println("Author is '" +
rstAuthors.getField("au_fname").getString() +" " +
rstAuthors.getField("au_lname").getString() +"'");
System.out.println("\nPress <Enter> to continue..");in.readLine();
}catch( AdoException ae )
{// Notify user of any errors that result from ADO.
// As passing a Recordset, check for null pointer first.if (rstAuthors != null)
{PrintProviderError(rstAuthors.getActiveConnection());
}else
{System.out.println("Exception: " + ae.getMessage());
}}
// System read requires this catch.catch( java.io.IOException je)
{PrintIOError(je);
}catch(java.lang.UnsatisfiedLinkError e)
{System.out.println("Exception: " + e.getMessage());
}catch(java.lang.NullPointerException ne)
{System.out.println(
"Exception: Attempt to use null where an object is required.");}
finally{
// Cleanup objects before exit.if (rstAuthors != null)
if (rstAuthors.getState() == 1)rstAuthors.close();
}}
// PrintProviderError Functionstatic void PrintProviderError( Connection Cnn1 )
{// Print Provider errors from Connection object.
// ErrItem is an item object in the Connection's Errors collection.com.ms.wfc.data.Error ErrItem = null;
long nCount = 0;int i = 0;
nCount = Cnn1.getErrors().getCount();// If there are any errors in the collection, print them.
if( nCount > 0);{
// Collection ranges from 0 to nCount - 1for (i = 0; i< nCount; i++)
{ErrItem = Cnn1.getErrors().getItem(i);
System.out.println("\t Error number: " + ErrItem.getNumber()+ "\t" + ErrItem.getDescription() );
}}
}// PrintIOError Function
static void PrintIOError( java.io.IOException je){
System.out.println("Error \n");System.out.println("\tSource = " + je.getClass() + "\n");
System.out.println("\tDescription = " + je.getMessage() + "\n");}
}
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]