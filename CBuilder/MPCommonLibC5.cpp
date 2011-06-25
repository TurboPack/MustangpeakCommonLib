//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
USERES("MPCommonLibC5.res");
USEPACKAGE("vcl50.bpi");
USEUNIT("..\Source\MPCommonObjects.pas");
USEUNIT("..\Source\MPCommonUtilities.pas");
USEUNIT("..\Source\MPDataObject.pas");
USEUNIT("..\Source\MPResources.pas");
USEUNIT("..\Source\MPShellTypes.pas");
USEUNIT("..\Source\MPShellUtilities.pas");
USEUNIT("..\Source\MPThreadManager.pas");
//---------------------------------------------------------------------------
#pragma package(smart_init)
//---------------------------------------------------------------------------

//   Package source.
//---------------------------------------------------------------------------

#pragma argsused
int WINAPI DllEntryPoint(HINSTANCE hinst, unsigned long reason, void*)
{
        return 1;
}
//---------------------------------------------------------------------------
