unit MPResources;

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

{$I Compilers.inc}
{$I ..\Include\Debug.inc}
{$I ..\Include\Addins.inc}

{$ifdef COMPILER_12_UP}
  {$WARN IMPLICIT_STRING_CAST       OFF}
 {$WARN IMPLICIT_STRING_CAST_LOSS  OFF}
{$endif COMPILER_12_UP}

{$B-}

{$IFDEF WEAKPACKAGING}
  // WARNING YOU CAN NOT REMOTE DEBUG A WEAK PACKAGE UNIT
  {$WEAKPACKAGEUNIT ON}
{$ENDIF}

const
  STR_LINKMANAGERDISABLED = 'ERROR: Accessing the Link Manager when it is not enabled';
  STR_NONEXISTENTCOLUMN = 'ERROR: Accessing a non-existent column index';
  STR_NONEXISTENTCOLUMNBANDROW = 'ERROR: Accessing a non-existent Column Band Row';
  STR_NONEXISTENTGROUP = 'ERROR: Accessing a non-existent Group ID';
  STR_UNREGISTEREDCONTROL = 'Error: Attemping to add a thread request for control or message id that is not registered';
  STR_ZEROWIDTHRECT = 'Error: Trying to fit text to a 0 width rectangle in the SplitTextW function';
  STR_BACKGROUNDALPHABLEND = 'Error: Background AlphaBlend mode requires an AlphaImage';

var
  // General Error message
  STR_ERROR: WideString = 'Error';

  // The name given a new folder when CreateNewFolder is called.
  STR_NEWFOLDER: WideString = 'New Folder';

  // The verb sent to the context menu notification events if the selected context
  // menu item is a non standard verb.
  STR_UNKNOWNCOMMAN: WideString = 'Unknown Command';

  STR_PROPERTIES: WideString = 'Properties';
  STR_PASTE: WideString = 'Paste';
  STR_PASTELINK: WideString = 'Paste Shortcut';

  // Names shown in column headers if toShellColumnDetails is not used. In that
  // case the shell handles the names based on local.
  STR_COLUMN_NAMES: array[0..9] of WideString = (
    'Name',
    'Size',
    'Type',
    'Modified',
    'Attributes',
    'Created',
    'Accessed',
    'Path',
    'DOS Name',
    'Custom'
  );

  // Strings that are used to show the attributes of a file in Details view.  Only
  // applies if toShellColumnDetails is not used.
  STR_ARCHIVE: WideString = 'A';
  STR_HIDDEN: WideString = 'H';
  STR_READONLY: WideString = 'R';
  STR_SYSTEM: WideString = 'S';
  STR_COMPRESS: WideString = 'C';

  // Strings that format the Details view in KB. Only applies if toShellColumnDetails
  // is not used.
  STR_FILE_SIZE_IN_KB: WideString = 'KB';
  STR_FILE_SIZE_IN_MB: WideString = 'MB';
  STR_FILE_SIZE_IN_TB: WideString = 'TB';
  STR_ZERO_KB: WideString = '0 KB';
  STR_ONE_KB: WideString = '1 KB';

  // What is displayed in the FileType column if VET could not get the information
  // the normal way and it had determined that the item is a system folder.
  STR_SYSTEMFOLDER: WideString = 'System Folder';
  STR_FILE: WideString = ' File'; // NT is lax in the FileType column if the file is not registered
                                  // it returns nothing.  This is tacked on the end of the file extension
                                  // for example 'PAS Files', BAK Files, ZIP Files and so on.

  // --------------------------------------------------------------------------
  // VirtualShellLink strings
  // --------------------------------------------------------------------------
  // Message shown if an attempt to create a new link is made with no target defined
  STR_NOTARGETDEFINED: WideString = 'No target application defined';
  // --------------------------------------------------------------------------

  // Shown when an operation in the TNamspace is being done on item that are not the
  // direct children of the TNamespace.  This is only a debugging aid, if the tree
  // is set up right this should never occur, i.e. restricted multiselect to one level
  STR_ERR_BAD_PIDL_RELATIONSHIP: WideString = 'PIDLs to operate on are not siblings of the Namespace doing the operation.';

implementation

end.
