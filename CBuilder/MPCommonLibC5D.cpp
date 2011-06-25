//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
USERES("MPCommonLibC5D.res");
USEPACKAGE("vcl50.bpi");
USEPACKAGE("MPCommonLibC5.bpi");
USEUNIT("..\Source\MPCommonWizardHelpers.pas");
USEUNIT("..\Source\MPCommonWizardTemplates.pas");
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
