---
title: InternetTimeout property example (VJ++)
ROBOTS: INDEX
ms.prod: access
ms.assetid: 7c09cd0b-b418-936f-766a-4cc14eea8e0b
ms.date: 06/08/2019
localization_priority: Normal
---


# InternetTimeout property example (VJ++)

**Applies to:** Access 2013 | Access 2016

This example demonstrates the [InternetTimeout](https://msdn.microsoft.com/library/66fc6e87-3d23-ce2c-18f5-0fc83ac43801%28Office.15%29.aspx) property, which exists on the [DataControl](https://msdn.microsoft.com/library/ac430669-7628-696c-c036-b5d35405d788%28Office.15%29.aspx) and [DataSpace](https://msdn.microsoft.com/library/7db181d5-422b-49fe-b6af-a20f5da520ff%28Office.15%29.aspx) objects. In this case, the **InternetTimout** property is demonstrated on the **DataControl** object and the timeout is set to 20 seconds.

```csharp

// BeginInternetTimeoutJ// The WFC class includes the ADO objects.
import com.ms.wfc.data.*;import com.ms.wfc.data.rds.*;
import java.io.* ; 
public class InternetTimeoutX{
// The main entry point for the application. 
public static void main (String[] args){
InternetTimeoutX();System.exit(0);
} 
// InternetTimeoutX function 
static void InternetTimeoutX(){ 
// Define ADO Objects.Recordset rstAuthors = null; 
// Declarations.BufferedReader in =
new BufferedReader (new InputStreamReader(System.in));int intCount = 0;
int intDisplaysize = 15; 
try{
IBindMgr dc = (IBindMgr) new DataControl();dc.setServer("https://MyServer");
dc.setConnect("DSN=pubs");dc.setSQL("SELECT * FROM Authors");
dc.setInternetTimeout(20000); // Wait at least 20 seconds.dc.Refresh();
rstAuthors = (Recordset)dc.getRecordset();while(!rstAuthors.getEOF())
{System.out.println(rstAuthors.getField
("au_fname").getString() + " " +rstAuthors.getField("au_lname").getString());
intCount++;if(intCount % intDisplaysize == 0)
{System.out.println("\nPress <Enter> to continue..");
in.readLine();intCount = 0;
}rstAuthors.moveNext();
} 
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
// PrintProviderError Function 
static void PrintProviderError( Connection Cnn1 ){
// Print Provider errors from Connection object.// ErrItem is an item object in the Connection's Errors collection.
com.ms.wfc.data.Error ErrItem = null;long nCount = 0;
int i = 0; 
nCount = Cnn1.getErrors().getCount(); 
// If there are any errors in the collection, print them.if( nCount > 0);
{// Collection ranges from 0 to nCount - 1
for (i = 0; i< nCount; i++){
ErrItem = Cnn1.getErrors().getItem(i);System.out.println("\t Error number: " + ErrItem.getNumber()
+ "\t" + ErrItem.getDescription() );}
} 
} 
// PrintIOError Function 
static void PrintIOError( java.io.IOException je){
System.out.println("Error \n");System.out.println("\tSource = " + je.getClass() + "\n");
System.out.println("\tDescription = " + je.getMessage() + "\n");}
}// EndInternetTimeoutJ
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]