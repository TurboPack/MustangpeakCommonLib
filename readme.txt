TurboPack MustangpeakCommonLib


Table of contents

1.  Introduction
2.  Package names
3.  Installation

==============================================


1. Introduction


TurboPack MustangpeakCommonLib is a base library needed for all Mustangpeak components.
This is a source-only release of TurboPack MustangpeakCommonLib. It includes
designtime and runtime packages for Delphi and CBuilder and supports 
Win32 and Win64.

==============================================

2. Package names


TurboPack MustangpeakCommonLib package names have the following form:

MPCommonLibD.bpl        (Delphi Runtime for all platforms)
MPCommonLibDD.bpl       (Delphi Runtime for the VCL)

MPCommonLibC.bpl        (C++Builder Runtime for all platforms)
MPCommonLibCD.bpl       (C++Builder Runtime for the VCL)


==============================================

3. Installation


To install TurboPack MustangpeakCommonLib into your IDE, take the following
steps:

  1. Unzip the release files into a directory (e.g., d:\MustangpeakCommonLib).

  2. Start RAD Studio.

  3. Add the source subdirectory (e.g., d:\MustangpeakCommonLib\source) to the
     IDE's library path. For C++Builder, add the hpp subdirectory
     (e.g., d:\MustangpeakCommonLib\source\hpp\Win32\Release) to the IDE's system include path.

  4. Open & install the designtime package specific to the IDE being
     used. The IDE should notify you the components have been
     installed.
