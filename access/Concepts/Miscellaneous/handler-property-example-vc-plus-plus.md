---
title: Handler property example (VC++)
ROBOTS: INDEX
ms.prod: access
ms.assetid: 9dcdb181-d4d9-36f9-ca64-153076af7205
ms.date: 06/08/2019
localization_priority: Normal
---


# Handler property example (VC++)

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

Your code looks like the following. The command assigned to the [SQL](sql-property-ado.md) property will match the **_AuthorById_** identifier and will retrieve a row for author Michael O'Leary. Although the [Connect](https://msdn.microsoft.com/library/11aa3284-18e9-6d2d-761b-c25090370b77%28Office.15%29.aspx) property in your code specifies the Northwind data source, that data source will be overwritten by the Msdfmap.ini _connect_ section. The **DataControl** object [Recordset](https://msdn.microsoft.com/library/5f4bb72d-ddfa-41c0-c353-b3a6632b4a91%28Office.15%29.aspx) property is assigned to a disconnected [Recordset](https://msdn.microsoft.com/library/0f963bf8-f066-dc8a-b754-f427de712df1%28Office.15%29.aspx) object purely as a coding convenience.

```cpp
// BeginHandlerCpp#import "c:\Program Files\Common Files\System\ADO\msado15.dll" \
no_namespace rename("EOF", "EndOfFile")#import "C:\Program Files\Common Files\System\MSADC\msadco.dll"
#include <ole2.h>#include <stdio.h>
#include <conio.h>// Function declarations
inline void TESTHR(HRESULT x) {if FAILED(x) _com_issue_error(x);};void HandlerX(void);
void PrintProviderError(_ConnectionPtr pConnection);void PrintComError(_com_error &e);
//////////////////////////////////////////////////////////// //
// Main Function //// //
//////////////////////////////////////////////////////////void main()
{HRESULT hr = S_OK;
hr = ::CoInitialize(NULL);if (SUCCEEDED(hr))
{HandlerX();
printf("Press any key to continue...");getch();
::CoUninitialize();}
}//////////////////////////////////////////////////////////
// //// HandlerX Function //
// ////////////////////////////////////////////////////////////
void HandlerX(void){
HRESULT hr = S_OK;// Define ADO object pointers.
// Initialize pointers on define.// These are in the ADODB:: namespace.
_RecordsetPtr pRst = NULL;//Define RDS object pointers.
RDS::IBindMgrPtr dc;try
{TESTHR(hr = dc.CreateInstance(__uuidof(RDS::DataControl)));
dc->Handler = "MSDFMAP.Handler";dc->Server = "https://MyServer";
dc->Connect = "Data Source=AuthorDatabase";dc->SQL = "AuthorById('267-41-2394')";
// Retrieve the record.dc->Refresh();
// Use another Recordset as a convenience.pRst = dc->GetRecordset();
printf("Author is %s %s",(LPSTR) (_bstr_t) pRst->Fields->GetItem("au_fname")->Value,\(LPSTR) (_bstr_t) pRst->Fields->GetItem("au_lname")->Value);
pRst->Close();} // End Try statement.
catch (_com_error &e){
PrintProviderError(pRst->GetActiveConnection());PrintComError(e);
}}
//////////////////////////////////////////////////////////// //
// PrintProviderError Function //// //
//////////////////////////////////////////////////////////void PrintProviderError(_ConnectionPtr pConnection)
{// Print Provider Errors from Connection object.
// pErr is a record object in the Connection's Error collection.ErrorPtr pErr = NULL;
long nCount = 0;long i = 0;
if( (pConnection->Errors->Count) > 0){
nCount = pConnection->Errors->Count;// Collection ranges from 0 to nCount -1.
for(i = 0; i < nCount; i++){
pErr = pConnection->Errors->GetItem(i);printf("\t Error number: %x\t%s", pErr->Number, pErr->Description);
}}
}//////////////////////////////////////////////////////////
//// PrintComError Function //
// ////////////////////////////////////////////////////////////
void PrintComError(_com_error &e){
_bstr_t bstrSource(e.Source());_bstr_t bstrDescription(e.Description());
// Print Com errors.printf("Error\n");
printf("\tCode = %08lx\n", e.Error());printf("\tCode meaning = %s\n", e.ErrorMessage());
printf("\tSource = %s\n", (LPCSTR) bstrSource);printf("\tDescription = %s\n", (LPCSTR) bstrDescription);
}// EndHandlerCpp
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]