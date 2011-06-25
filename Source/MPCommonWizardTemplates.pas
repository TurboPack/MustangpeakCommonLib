unit MPCommonWizardTemplates;

// Version 2.1.0
//
// The contents of this file are subject to the Mozilla Public License
// Version 1.1 (the "License"); you may not use this file except in compliance
// with the License. You may obtain a copy of the License at http://www.mozilla.org/MPL/
//
// Alternatively, you may redistribute this library, use and/or modify it under the terms of the
// GNU Lesser General Public License as published by the Free Software Foundation;
// either version 2.1 of the License, or (at your option) any later version.
// You may obtain a copy of the LGPL at http://www.gnu.org/copyleft/.
//
// Software distributed under the License is distributed on an "AS IS" basis,
// WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License for the
// specific language governing rights and limitations under the License.
//
// The initial developer of this code is Jim Kueneman <jimdk@mindspring.com>
// Special thanks to the following in no particular order for their help/support/code
//    Danijel Malik, Robert Lee, Werner Lehmann, Alexey Torgashin, Milan Vandrovec
//
//----------------------------------------------------------------------------

interface

{$I ..\Include\Addins.inc}

{$ifdef COMPILER_12_UP}
  {$WARN IMPLICIT_STRING_CAST       OFF}
 {$WARN IMPLICIT_STRING_CAST_LOSS  OFF}
{$endif COMPILER_12_UP}

const
  FILE_FORM_TEMPLATE_DELPHI =

  'unit %ModuleIdent;                                                ' + #13#10 +
  '                                                                  ' + #13#10 +
  'interface                                                         ' + #13#10 +
  '                                                                  ' + #13#10 +
  '                                                                  ' + #13#10 +
  'uses                                                              ' + #13#10 +
  '  %IncludeList                                                    ' + #13#10 +
  '                                                                  ' + #13#10 +
  'type                                                              ' + #13#10 +
  '  T%FormIdent= class(T%AncestorIdent)                             ' + #13#10 +
  '  private                                                         ' + #13#10 +
  '    { Private declarations }                                      ' + #13#10 +
  '  public                                                          ' + #13#10 +
  '    { Public declarations }                                       ' + #13#10 +
  '  end;                                                            ' + #13#10 +
  '                                                                  ' + #13#10 +
  'var                                                               ' + #13#10 +
  '  %FormIdent: T%FormIdent;                                        ' + #13#10 +
  '                                                                  ' + #13#10 +
  'implementation                                                    ' + #13#10 +
   '                                                                 ' + #13#10 +
  '{$R *.dfm}                                                        ' + #13#10 +
  '                                                                  ' + #13#10 +
  '                                                                  ' + #13#10 +
  'end.                                                              ';

  FILE_FORM_TEMPLATE_BCB =
  '//----------------------------------------------------------------' + #13#10 +
  '%IncludeList                                                      ' + #13#10 +
  '#pragma hdrstop                                                   ' + #13#10 +
  '                                                                  ' + #13#10 +
  '#include "%ModuleIdent.h"                                         ' + #13#10 +
  '//----------------------------------------------------------------' + #13#10 +
  '#pragma package(smart_init)                                       ' + #13#10 +
  '#pragma resource "*.dfm"                                          ' + #13#10 +
  'T%FormIdent *%FormIdent;                                          ' + #13#10 +
  '//----------------------------------------------------------------' + #13#10 +
  '__fastcall T%FormIdent::T%FormIdent(TComponent* Owner)            ' + #13#10 +
  '    : T%AncestorIdent(Owner)                                      ' + #13#10 +
  '{                                                                 ' + #13#10 +
  '}                                                                 ' + #13#10 +
  '//----------------------------------------------------------------';

implementation

end.
