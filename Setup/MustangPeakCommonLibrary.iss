
[Setup]
AppName=Mustangpeak Common Library
; (use an include file for the AppVerName key for dynamic version number creation.)
#include "Version.txt"
DefaultDirName={commondocs}\Mustangpeak\Common Library
DefaultGroupName=MustangPeakCommonLib
LicenseFile=..\Docs\Licence.txt
UsePreviousAppDir=true
UsePreviousGroup=true
ShowLanguageDialog=yes
OutputDir=.\

OutputBaseFilename=MustangPeakCommonLib
AppCopyright=©Jim Kueneman, Mustangpeak
AllowNoIcons=true

WizardSmallImageFile=..\..\InnoSetup\Shared\SantaRitasSmall.bmp
;WizardImageFile=D:\Program Data\Delphi Projects\3rd Party Components\InnoSetup\Shared\mustangpeak.bmp
WizardImageStretch=false

[Messages]
BeveledLabel=© Jim Kueneman, Mustangpeak.net

[Files]
Source: Setup.ini; DestDir: {app}
Source: ..\..\InnoSetup\MustangpeakComponentInstaller.dll; DestDir: {app}
Source: ..\CBuilder\MPCommonLibC6.bpk; DestDir: {app}\CBuilder\
Source: ..\CBuilder\MPCommonLibC6.cpp; DestDir: {app}\CBuilder\
Source: ..\CBuilder\MPCommonLibC6.res; DestDir: {app}\CBuilder\
Source: ..\CBuilder\MPCommonLibC5.bpk; DestDir: {app}\CBuilder\
Source: ..\CBuilder\MPCommonLibC5.cpp; DestDir: {app}\CBuilder\
Source: ..\CBuilder\MPCommonLibC5.res; DestDir: {app}\CBuilder\
Source: ..\CBuilder\MPCommonLibC6D.bpk; DestDir: {app}\CBuilder\
Source: ..\CBuilder\MPCommonLibC6D.cpp; DestDir: {app}\CBuilder\
Source: ..\CBuilder\MPCommonLibC6D.res; DestDir: {app}\CBuilder\
Source: ..\CBuilder\MPCommonLibC5D.bpk; DestDir: {app}\CBuilder\
Source: ..\CBuilder\MPCommonLibC5D.cpp; DestDir: {app}\CBuilder\
Source: ..\CBuilder\MPCommonLibC5D.res; DestDir: {app}\CBuilder\
Source: ..\Delphi\MPCommonLibD4.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD4.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD5.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD5.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD6.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD6.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD7.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD7.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD9.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD9.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD10.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD10.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD11.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD11.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD12.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD12.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD14.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD14.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD15.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD15.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD4D.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD4D.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD5D.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD5D.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD6D.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD6D.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD7D.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD7D.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD9D.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD9D.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD10D.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD10D.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD11D.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD11D.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD12D.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD12D.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD14D.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD14D.res; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD15D.dpk; DestDir: {app}\Delphi\
Source: ..\Delphi\MPCommonLibD15D.res; DestDir: {app}\Delphi\
Source: ..\Docs\Licence.txt; DestDir: {app}\Docs\
Source: ..\Docs\Whats New.txt; DestDir: {app}\Docs\
Source: ..\Docs\Unicode Compatibility.txt; DestDir: {app}\Docs\
Source: ..\Include\Debug.inc; DestDir: {app}\Include\
Source: ..\Include\AddIns.inc; DestDir: {app}\Include\
Source: ..\Source\Compilers.inc; DestDir: {app}\Source\
Source: ..\Source\MPCommonObjects.pas; DestDir: {app}\Source\
Source: ..\Source\MPCommonUtilities.pas; DestDir: {app}\Source\
Source: ..\Source\MPDataObject.pas; DestDir: {app}\Source\
Source: ..\Source\MPResources.pas; DestDir: {app}\Source\
Source: ..\Source\MPShellTypes.pas; DestDir: {app}\Source\
Source: ..\Source\MPThreadManager.pas; DestDir: {app}\Source\
Source: ..\Source\MPShellUtilities.pas; DestDir: {app}\Source\
Source: ..\Source\MPCommonWizardTemplates.pas; DestDir: {app}\Source\
Source: ..\Source\MPCommonWizardHelpers.pas; DestDir: {app}\Source\
Source: ..\Source\Options.inc; DestDir: {app}\Source\

[Registry]
Root: HKCU; Subkey: Software\Mustangpeak\CommonLib; ValueType: string; ValueName: InstallPath; ValueData: {app}; Flags: uninsdeletekey
Root: HKCU; Subkey: Software\Mustangpeak\CommonLib; ValueType: string; ValueName: SourcePath; ValueData: {app}\Source\; Flags: uninsdeletekey
Root: HKCU; Subkey: Software\Mustangpeak\CommonLib; ValueType: string; ValueName: DelphiPackagePath; ValueData: {app}\Delphi\; Flags: uninsdeletekey
Root: HKCU; Subkey: Software\Mustangpeak\CommonLib; ValueType: string; ValueName: CBuilderPackagePath; ValueData: {app}\CBuilder\; Flags: uninsdeletekey

[Code]
const
  SetupFile = 'Setup.ini';
  WaitForIDEs = False;

#include "..\..\InnoSetup\Templates\IDE Install.txt"
[Dirs]
Name: {app}\Source
Name: {app}\Delphi
Name: {app}\Docs
Name: {app}\Include

Name: {app}\CBuilder
Name: {app}\Delphi
