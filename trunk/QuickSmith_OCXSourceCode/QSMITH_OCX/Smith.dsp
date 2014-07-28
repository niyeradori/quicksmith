# Microsoft Developer Studio Project File - Name="Smith" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 5.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=Smith - Win32 Release
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "Smith.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "Smith.mak" CFG="Smith - Win32 Release"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "Smith - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "Smith - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "Smith - Win32 Unicode Debug" (based on\
 "Win32 (x86) Dynamic-Link Library")
!MESSAGE "Smith - Win32 Unicode Release" (based on\
 "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "Smith - Win32 Release"

# PROP BASE Use_MFC 2
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir ".\Release"
# PROP BASE Intermediate_Dir ".\Release"
# PROP BASE Target_Ext "ocx"
# PROP BASE Target_Dir ""
# PROP Use_MFC 2
# PROP Use_Debug_Libraries 0
# PROP Output_Dir ".\Release"
# PROP Intermediate_Dir ".\Release"
# PROP Target_Ext "ocx"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /Yu"stdafx.h" /c
# ADD CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /Yu"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 /nologo /subsystem:windows /dll /machine:I386
# Begin Custom Build - Registering OLE control...
OutDir=.\.\Release
TargetPath=.\Release\Smith.ocx
InputPath=.\Release\Smith.ocx
SOURCE=$(InputPath)

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	
# End Custom Build

!ELSEIF  "$(CFG)" == "Smith - Win32 Debug"

# PROP BASE Use_MFC 2
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir ".\Debug"
# PROP BASE Intermediate_Dir ".\Debug"
# PROP BASE Target_Ext "ocx"
# PROP BASE Target_Dir ""
# PROP Use_MFC 2
# PROP Use_Debug_Libraries 1
# PROP Output_Dir ".\Debug"
# PROP Intermediate_Dir ".\Debug"
# PROP Target_Ext "ocx"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MDd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /Yu"stdafx.h" /c
# ADD CPP /nologo /MDd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /Yu"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "_DEBUG" /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /debug /machine:I386
# ADD LINK32 /nologo /subsystem:windows /dll /debug /machine:I386
# Begin Custom Build - Registering OLE control...
OutDir=.\.\Debug
TargetPath=.\Debug\Smith.ocx
InputPath=.\Debug\Smith.ocx
SOURCE=$(InputPath)

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	
# End Custom Build

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Debug"

# PROP BASE Use_MFC 2
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir ".\DebugU"
# PROP BASE Intermediate_Dir ".\DebugU"
# PROP BASE Target_Ext "ocx"
# PROP BASE Target_Dir ""
# PROP Use_MFC 2
# PROP Use_Debug_Libraries 1
# PROP Output_Dir ".\UniDebug"
# PROP Intermediate_Dir ".\UniDebug"
# PROP Target_Ext "ocx"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MDd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /Yu"stdafx.h" /c
# ADD CPP /nologo /MDd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_UNICODE" /Yu"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "_DEBUG" /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /debug /machine:I386
# ADD LINK32 /nologo /subsystem:windows /dll /debug /machine:I386
# Begin Custom Build - Registering OLE control...
OutDir=.\.\UniDebug
TargetPath=.\UniDebug\Smith.ocx
InputPath=.\UniDebug\Smith.ocx
SOURCE=$(InputPath)

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	
# End Custom Build

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Release"

# PROP BASE Use_MFC 2
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir ".\ReleaseU"
# PROP BASE Intermediate_Dir ".\ReleaseU"
# PROP BASE Target_Ext "ocx"
# PROP BASE Target_Dir ""
# PROP Use_MFC 2
# PROP Use_Debug_Libraries 0
# PROP Output_Dir ".\UniRelease"
# PROP Intermediate_Dir ".\UniRelease"
# PROP Target_Ext "ocx"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /Yu"stdafx.h" /c
# ADD CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_UNICODE" /Yu"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 /nologo /subsystem:windows /dll /machine:I386
# Begin Custom Build - Registering OLE control...
OutDir=.\.\UniRelease
TargetPath=.\UniRelease\Smith.ocx
InputPath=.\UniRelease\Smith.ocx
SOURCE=$(InputPath)

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	
# End Custom Build

!ENDIF 

# Begin Target

# Name "Smith - Win32 Release"
# Name "Smith - Win32 Debug"
# Name "Smith - Win32 Unicode Debug"
# Name "Smith - Win32 Unicode Release"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;hpj;bat;for;f90"
# Begin Source File

SOURCE=.\Datappg.cpp

!IF  "$(CFG)" == "Smith - Win32 Release"

# ADD CPP /W4

!ELSEIF  "$(CFG)" == "Smith - Win32 Debug"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Debug"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\readme.txt

!IF  "$(CFG)" == "Smith - Win32 Release"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "Smith - Win32 Debug"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Debug"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Smith.cpp

!IF  "$(CFG)" == "Smith - Win32 Release"

# ADD CPP /W4

!ELSEIF  "$(CFG)" == "Smith - Win32 Debug"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Debug"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Smith.def
# End Source File
# Begin Source File

SOURCE=.\Smith.odl
# End Source File
# Begin Source File

SOURCE=.\Smith.rc

!IF  "$(CFG)" == "Smith - Win32 Release"

!ELSEIF  "$(CFG)" == "Smith - Win32 Debug"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Debug"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Smithctl.cpp

!IF  "$(CFG)" == "Smith - Win32 Release"

# ADD CPP /W4

!ELSEIF  "$(CFG)" == "Smith - Win32 Debug"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Debug"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Smithppg.cpp

!IF  "$(CFG)" == "Smith - Win32 Release"

# ADD CPP /W4

!ELSEIF  "$(CFG)" == "Smith - Win32 Debug"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Debug"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\stdafx.cpp

!IF  "$(CFG)" == "Smith - Win32 Release"

# ADD CPP /W4 /Yc"stdafx.h"

!ELSEIF  "$(CFG)" == "Smith - Win32 Debug"

# ADD CPP /Yc"stdafx.h"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Debug"

# ADD CPP /Yc"stdafx.h"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Release"

# ADD CPP /Yc"stdafx.h"

!ENDIF 

# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl;fi;fd"
# Begin Source File

SOURCE=.\Datappg.h
# End Source File
# Begin Source File

SOURCE=.\Resource.h
# End Source File
# Begin Source File

SOURCE=.\Smith.h
# End Source File
# Begin Source File

SOURCE=.\smithctl.h
# End Source File
# Begin Source File

SOURCE=.\Smithppg.h
# End Source File
# Begin Source File

SOURCE=.\stdafx.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;cnt;rtf;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\Smith.ico
# End Source File
# Begin Source File

SOURCE=.\Smith.rc2
# End Source File
# Begin Source File

SOURCE=.\Smithctl.bmp
# End Source File
# End Group
# End Target
# End Project
