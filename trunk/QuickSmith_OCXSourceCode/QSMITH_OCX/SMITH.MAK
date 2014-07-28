# Microsoft Developer Studio Generated NMAKE File, Format Version 4.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

!IF "$(CFG)" == ""
CFG=Smith - Win32 Debug
!MESSAGE No configuration specified.  Defaulting to Smith - Win32 Debug.
!ENDIF 

!IF "$(CFG)" != "Smith - Win32 Release" && "$(CFG)" != "Smith - Win32 Debug" &&\
 "$(CFG)" != "Smith - Win32 Unicode Debug" && "$(CFG)" !=\
 "Smith - Win32 Unicode Release"
!MESSAGE Invalid configuration "$(CFG)" specified.
!MESSAGE You can specify a configuration when running NMAKE on this makefile
!MESSAGE by defining the macro CFG on the command line.  For example:
!MESSAGE 
!MESSAGE NMAKE /f "smith.mak" CFG="Smith - Win32 Debug"
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
!ERROR An invalid configuration is specified.
!ENDIF 

!IF "$(OS)" == "Windows_NT"
NULL=
!ELSE 
NULL=nul
!ENDIF 
################################################################################
# Begin Project
# PROP Target_Last_Scanned "Smith - Win32 Debug"
CPP=cl.exe
RSC=rc.exe
MTL=mktyplib.exe

!IF  "$(CFG)" == "Smith - Win32 Release"

# PROP BASE Use_MFC 2
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP BASE Target_Ext "ocx"
# PROP Use_MFC 2
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release"
# PROP Target_Dir ""
# PROP Target_Ext "ocx"
OUTDIR=.\Release
INTDIR=.\Release
# Begin Custom Macros
OutDir=.\Release
# End Custom Macros

ALL : "$(OUTDIR)\smith.ocx" "$(OUTDIR)\regsvr32.trg"

CLEAN : 
	-@erase ".\Release\smith.lib"
	-@erase ".\Release\Smithppg.obj"
	-@erase ".\Release\smith.pch"
	-@erase ".\Release\Smith.obj"
	-@erase ".\Release\Smithctl.obj"
	-@erase ".\Release\stdafx.obj"
	-@erase ".\Release\Smith.res"
	-@erase ".\Release\Smith.tlb"
	-@erase ".\Release\smith.exp"
	-@erase ".\Release\regsvr32.trg"
	-@erase ".\Release\smith.ocx"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

# ADD BASE CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /Yu"stdafx.h" /c
# ADD CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /Yu"stdafx.h" /c
CPP_PROJ=/nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D\
 "_AFXDLL" /D "_MBCS" /Fp"$(INTDIR)/smith.pch" /Yu"stdafx.h" /Fo"$(INTDIR)/"\
 /c 
CPP_OBJS=.\Release/
CPP_SBRS=
# ADD BASE MTL /nologo /D "NDEBUG" /win32
# ADD MTL /nologo /D "NDEBUG" /win32
MTL_PROJ=/nologo /D "NDEBUG" /win32 
# ADD BASE RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
RSC_PROJ=/l 0x409 /fo"$(INTDIR)/Smith.res" /d "NDEBUG" /d "_AFXDLL" 
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o"$(OUTDIR)/smith.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 /nologo /subsystem:windows /dll /machine:I386
LINK32_FLAGS=/nologo /subsystem:windows /dll /incremental:no\
 /pdb:"$(OUTDIR)/smith.pdb" /machine:I386 /def:".\Smith.def"\
 /out:"$(OUTDIR)/smith.ocx" /implib:"$(OUTDIR)/smith.lib" 
DEF_FILE= \
	".\Smith.def"
LINK32_OBJS= \
	"$(INTDIR)/Smithppg.obj" \
	"$(INTDIR)/Smith.obj" \
	"$(INTDIR)/Smithctl.obj" \
	"$(INTDIR)/stdafx.obj" \
	"$(INTDIR)/Smith.res"

"$(OUTDIR)\smith.ocx" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

# Begin Custom Build - Registering OLE control...
OutDir=.\Release
TargetPath=.\Release\smith.ocx
InputPath=.\Release\smith.ocx
SOURCE=$(InputPath)

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   regsvr32 /s /c "$(TargetPath)"
   echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg"

# End Custom Build

!ELSEIF  "$(CFG)" == "Smith - Win32 Debug"

# PROP BASE Use_MFC 2
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP BASE Target_Ext "ocx"
# PROP Use_MFC 2
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Target_Dir ""
# PROP Target_Ext "ocx"
OUTDIR=.\Debug
INTDIR=.\Debug
# Begin Custom Macros
OutDir=.\Debug
# End Custom Macros

ALL : "$(OUTDIR)\smith.ocx" "$(OUTDIR)\regsvr32.trg"

CLEAN : 
	-@erase ".\Debug\vc40.pdb"
	-@erase ".\Debug\smith.pch"
	-@erase ".\Debug\vc40.idb"
	-@erase ".\Debug\smith.ilk"
	-@erase ".\Debug\Smith.obj"
	-@erase ".\Debug\stdafx.obj"
	-@erase ".\Debug\Smithppg.obj"
	-@erase ".\Debug\Smithctl.obj"
	-@erase ".\Debug\Smith.res"
	-@erase ".\Debug\smith.tlb"
	-@erase ".\Debug\smith.lib"
	-@erase ".\Debug\smith.exp"
	-@erase ".\Debug\smith.pdb"
	-@erase ".\Debug\regsvr32.trg"
	-@erase ".\Debug\smith.ocx"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

# ADD BASE CPP /nologo /MDd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /Yu"stdafx.h" /c
# ADD CPP /nologo /MDd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /Yu"stdafx.h" /c
CPP_PROJ=/nologo /MDd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS"\
 /D "_AFXDLL" /D "_MBCS" /Fp"$(INTDIR)/smith.pch" /Yu"stdafx.h"\
 /Fo"$(INTDIR)/" /Fd"$(INTDIR)/" /c 
CPP_OBJS=.\Debug/
CPP_SBRS=
# ADD BASE MTL /nologo /D "_DEBUG" /win32
# ADD MTL /nologo /D "_DEBUG" /win32
MTL_PROJ=/nologo /D "_DEBUG" /win32 
# ADD BASE RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
RSC_PROJ=/l 0x409 /fo"$(INTDIR)/Smith.res" /d "_DEBUG" /d "_AFXDLL" 
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o"$(OUTDIR)/smith.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /debug /machine:I386
# ADD LINK32 /nologo /subsystem:windows /dll /debug /machine:I386
LINK32_FLAGS=/nologo /subsystem:windows /dll /incremental:yes\
 /pdb:"$(OUTDIR)/smith.pdb" /debug /machine:I386 /def:".\Smith.def"\
 /out:"$(OUTDIR)/smith.ocx" /implib:"$(OUTDIR)/smith.lib" 
DEF_FILE= \
	".\Smith.def"
LINK32_OBJS= \
	"$(INTDIR)/Smith.obj" \
	"$(INTDIR)/stdafx.obj" \
	"$(INTDIR)/Smithppg.obj" \
	"$(INTDIR)/Smithctl.obj" \
	"$(INTDIR)/Smith.res"

"$(OUTDIR)\smith.ocx" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

# Begin Custom Build - Registering OLE control...
OutDir=.\Debug
TargetPath=.\Debug\smith.ocx
InputPath=.\Debug\smith.ocx
SOURCE=$(InputPath)

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   regsvr32 /s /c "$(TargetPath)"
   echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg"

# End Custom Build

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Debug"

# PROP BASE Use_MFC 2
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "DebugU"
# PROP BASE Intermediate_Dir "DebugU"
# PROP BASE Target_Dir ""
# PROP BASE Target_Ext "ocx"
# PROP Use_MFC 2
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "UniDebug"
# PROP Intermediate_Dir "UniDebug"
# PROP Target_Dir ""
# PROP Target_Ext "ocx"
OUTDIR=.\UniDebug
INTDIR=.\UniDebug
# Begin Custom Macros
OutDir=.\UniDebug
# End Custom Macros

ALL : "$(OUTDIR)\smith.ocx" "$(OUTDIR)\regsvr32.trg"

CLEAN : 
	-@erase ".\UniDebug\vc40.pdb"
	-@erase ".\UniDebug\smith.pch"
	-@erase ".\UniDebug\vc40.idb"
	-@erase ".\UniDebug\smith.ilk"
	-@erase ".\UniDebug\Smithctl.obj"
	-@erase ".\UniDebug\Smith.obj"
	-@erase ".\UniDebug\Smithppg.obj"
	-@erase ".\UniDebug\stdafx.obj"
	-@erase ".\UniDebug\Smith.res"
	-@erase ".\UniDebug\Smith.tlb"
	-@erase ".\UniDebug\smith.lib"
	-@erase ".\UniDebug\smith.exp"
	-@erase ".\UniDebug\smith.pdb"
	-@erase ".\UniDebug\regsvr32.trg"
	-@erase ".\UniDebug\smith.ocx"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

# ADD BASE CPP /nologo /MDd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /Yu"stdafx.h" /c
# ADD CPP /nologo /MDd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_UNICODE" /Yu"stdafx.h" /c
CPP_PROJ=/nologo /MDd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS"\
 /D "_AFXDLL" /D "_UNICODE" /Fp"$(INTDIR)/smith.pch" /Yu"stdafx.h"\
 /Fo"$(INTDIR)/" /Fd"$(INTDIR)/" /c 
CPP_OBJS=.\UniDebug/
CPP_SBRS=
# ADD BASE MTL /nologo /D "_DEBUG" /win32
# ADD MTL /nologo /D "_DEBUG" /win32
MTL_PROJ=/nologo /D "_DEBUG" /win32 
# ADD BASE RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
RSC_PROJ=/l 0x409 /fo"$(INTDIR)/Smith.res" /d "_DEBUG" /d "_AFXDLL" 
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o"$(OUTDIR)/smith.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /debug /machine:I386
# ADD LINK32 /nologo /subsystem:windows /dll /debug /machine:I386
LINK32_FLAGS=/nologo /subsystem:windows /dll /incremental:yes\
 /pdb:"$(OUTDIR)/smith.pdb" /debug /machine:I386 /def:".\Smith.def"\
 /out:"$(OUTDIR)/smith.ocx" /implib:"$(OUTDIR)/smith.lib" 
DEF_FILE= \
	".\Smith.def"
LINK32_OBJS= \
	"$(INTDIR)/Smithctl.obj" \
	"$(INTDIR)/Smith.obj" \
	"$(INTDIR)/Smithppg.obj" \
	"$(INTDIR)/stdafx.obj" \
	"$(INTDIR)/Smith.res"

"$(OUTDIR)\smith.ocx" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

# Begin Custom Build - Registering OLE control...
OutDir=.\UniDebug
TargetPath=.\UniDebug\smith.ocx
InputPath=.\UniDebug\smith.ocx
SOURCE=$(InputPath)

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   regsvr32 /s /c "$(TargetPath)"
   echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg"

# End Custom Build

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Release"

# PROP BASE Use_MFC 2
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "ReleaseU"
# PROP BASE Intermediate_Dir "ReleaseU"
# PROP BASE Target_Dir ""
# PROP BASE Target_Ext "ocx"
# PROP Use_MFC 2
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "UniRelease"
# PROP Intermediate_Dir "UniRelease"
# PROP Target_Dir ""
# PROP Target_Ext "ocx"
OUTDIR=.\UniRelease
INTDIR=.\UniRelease
# Begin Custom Macros
OutDir=.\UniRelease
# End Custom Macros

ALL : "$(OUTDIR)\smith.ocx" "$(OUTDIR)\regsvr32.trg"

CLEAN : 
	-@erase ".\UniRelease\smith.lib"
	-@erase ".\UniRelease\stdafx.obj"
	-@erase ".\UniRelease\Smithppg.obj"
	-@erase ".\UniRelease\smith.pch"
	-@erase ".\UniRelease\Smith.obj"
	-@erase ".\UniRelease\Smithctl.obj"
	-@erase ".\UniRelease\Smith.res"
	-@erase ".\UniRelease\Smith.tlb"
	-@erase ".\UniRelease\smith.exp"
	-@erase ".\UniRelease\regsvr32.trg"
	-@erase ".\UniRelease\smith.ocx"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

# ADD BASE CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /Yu"stdafx.h" /c
# ADD CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_UNICODE" /Yu"stdafx.h" /c
CPP_PROJ=/nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D\
 "_AFXDLL" /D "_UNICODE" /Fp"$(INTDIR)/smith.pch" /Yu"stdafx.h"\
 /Fo"$(INTDIR)/" /c 
CPP_OBJS=.\UniRelease/
CPP_SBRS=
# ADD BASE MTL /nologo /D "NDEBUG" /win32
# ADD MTL /nologo /D "NDEBUG" /win32
MTL_PROJ=/nologo /D "NDEBUG" /win32 
# ADD BASE RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
RSC_PROJ=/l 0x409 /fo"$(INTDIR)/Smith.res" /d "NDEBUG" /d "_AFXDLL" 
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o"$(OUTDIR)/smith.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 /nologo /subsystem:windows /dll /machine:I386
LINK32_FLAGS=/nologo /subsystem:windows /dll /incremental:no\
 /pdb:"$(OUTDIR)/smith.pdb" /machine:I386 /def:".\Smith.def"\
 /out:"$(OUTDIR)/smith.ocx" /implib:"$(OUTDIR)/smith.lib" 
DEF_FILE= \
	".\Smith.def"
LINK32_OBJS= \
	"$(INTDIR)/stdafx.obj" \
	"$(INTDIR)/Smithppg.obj" \
	"$(INTDIR)/Smith.obj" \
	"$(INTDIR)/Smithctl.obj" \
	"$(INTDIR)/Smith.res"

"$(OUTDIR)\smith.ocx" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

# Begin Custom Build - Registering OLE control...
OutDir=.\UniRelease
TargetPath=.\UniRelease\smith.ocx
InputPath=.\UniRelease\smith.ocx
SOURCE=$(InputPath)

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   regsvr32 /s /c "$(TargetPath)"
   echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg"

# End Custom Build

!ENDIF 

.c{$(CPP_OBJS)}.obj:
   $(CPP) $(CPP_PROJ) $<  

.cpp{$(CPP_OBJS)}.obj:
   $(CPP) $(CPP_PROJ) $<  

.cxx{$(CPP_OBJS)}.obj:
   $(CPP) $(CPP_PROJ) $<  

.c{$(CPP_SBRS)}.sbr:
   $(CPP) $(CPP_PROJ) $<  

.cpp{$(CPP_SBRS)}.sbr:
   $(CPP) $(CPP_PROJ) $<  

.cxx{$(CPP_SBRS)}.sbr:
   $(CPP) $(CPP_PROJ) $<  

################################################################################
# Begin Target

# Name "Smith - Win32 Release"
# Name "Smith - Win32 Debug"
# Name "Smith - Win32 Unicode Debug"
# Name "Smith - Win32 Unicode Release"

!IF  "$(CFG)" == "Smith - Win32 Release"

!ELSEIF  "$(CFG)" == "Smith - Win32 Debug"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Debug"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Release"

!ENDIF 

################################################################################
# Begin Source File

SOURCE=.\readme.txt

!IF  "$(CFG)" == "Smith - Win32 Release"

!ELSEIF  "$(CFG)" == "Smith - Win32 Debug"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Debug"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Release"

!ENDIF 

# End Source File
################################################################################
# Begin Source File

SOURCE=.\stdafx.cpp
DEP_CPP_STDAF=\
	".\stdafx.h"\
	

!IF  "$(CFG)" == "Smith - Win32 Release"

# ADD CPP /Yc"stdafx.h"

BuildCmds= \
	$(CPP) /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D\
 "_AFXDLL" /D "_MBCS" /Fp"$(INTDIR)/smith.pch" /Yc"stdafx.h" /Fo"$(INTDIR)/"\
 /c $(SOURCE) \
	

"$(INTDIR)\stdafx.obj" : $(SOURCE) $(DEP_CPP_STDAF) "$(INTDIR)"
   $(BuildCmds)

"$(INTDIR)\smith.pch" : $(SOURCE) $(DEP_CPP_STDAF) "$(INTDIR)"
   $(BuildCmds)

!ELSEIF  "$(CFG)" == "Smith - Win32 Debug"

# ADD CPP /Yc"stdafx.h"

BuildCmds= \
	$(CPP) /nologo /MDd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS"\
 /D "_AFXDLL" /D "_MBCS" /Fp"$(INTDIR)/smith.pch" /Yc"stdafx.h"\
 /Fo"$(INTDIR)/" /Fd"$(INTDIR)/" /c $(SOURCE) \
	

"$(INTDIR)\stdafx.obj" : $(SOURCE) $(DEP_CPP_STDAF) "$(INTDIR)"
   $(BuildCmds)

"$(INTDIR)\smith.pch" : $(SOURCE) $(DEP_CPP_STDAF) "$(INTDIR)"
   $(BuildCmds)

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Debug"

# ADD CPP /Yc"stdafx.h"

BuildCmds= \
	$(CPP) /nologo /MDd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS"\
 /D "_AFXDLL" /D "_UNICODE" /Fp"$(INTDIR)/smith.pch" /Yc"stdafx.h"\
 /Fo"$(INTDIR)/" /Fd"$(INTDIR)/" /c $(SOURCE) \
	

"$(INTDIR)\stdafx.obj" : $(SOURCE) $(DEP_CPP_STDAF) "$(INTDIR)"
   $(BuildCmds)

"$(INTDIR)\smith.pch" : $(SOURCE) $(DEP_CPP_STDAF) "$(INTDIR)"
   $(BuildCmds)

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Release"

# ADD CPP /Yc"stdafx.h"

BuildCmds= \
	$(CPP) /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D\
 "_AFXDLL" /D "_UNICODE" /Fp"$(INTDIR)/smith.pch" /Yc"stdafx.h"\
 /Fo"$(INTDIR)/" /c $(SOURCE) \
	

"$(INTDIR)\stdafx.obj" : $(SOURCE) $(DEP_CPP_STDAF) "$(INTDIR)"
   $(BuildCmds)

"$(INTDIR)\smith.pch" : $(SOURCE) $(DEP_CPP_STDAF) "$(INTDIR)"
   $(BuildCmds)

!ENDIF 

# End Source File
################################################################################
# Begin Source File

SOURCE=.\Smith.cpp

!IF  "$(CFG)" == "Smith - Win32 Release"

DEP_CPP_SMITH=\
	".\Smith.h"\
	".\stdafx.h"\
	

"$(INTDIR)\Smith.obj" : $(SOURCE) $(DEP_CPP_SMITH) "$(INTDIR)"\
 "$(INTDIR)\smith.pch"


!ELSEIF  "$(CFG)" == "Smith - Win32 Debug"

DEP_CPP_SMITH=\
	".\stdafx.h"\
	".\Smith.h"\
	

"$(INTDIR)\Smith.obj" : $(SOURCE) $(DEP_CPP_SMITH) "$(INTDIR)"\
 "$(INTDIR)\smith.pch"


!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Debug"

DEP_CPP_SMITH=\
	".\Smith.h"\
	".\stdafx.h"\
	

"$(INTDIR)\Smith.obj" : $(SOURCE) $(DEP_CPP_SMITH) "$(INTDIR)"\
 "$(INTDIR)\smith.pch"


!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Release"

DEP_CPP_SMITH=\
	".\Smith.h"\
	".\stdafx.h"\
	

"$(INTDIR)\Smith.obj" : $(SOURCE) $(DEP_CPP_SMITH) "$(INTDIR)"\
 "$(INTDIR)\smith.pch"


!ENDIF 

# End Source File
################################################################################
# Begin Source File

SOURCE=.\Smith.def

!IF  "$(CFG)" == "Smith - Win32 Release"

!ELSEIF  "$(CFG)" == "Smith - Win32 Debug"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Debug"

!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Release"

!ENDIF 

# End Source File
################################################################################
# Begin Source File

SOURCE=.\Smith.rc

!IF  "$(CFG)" == "Smith - Win32 Release"

DEP_RSC_SMITH_=\
	".\Smith.ico"\
	".\Smithctl.bmp"\
	".\Resource.h"\
	".\Smith.rc2"\
	".\Release\Smith.tlb"\
	

"$(INTDIR)\Smith.res" : $(SOURCE) $(DEP_RSC_SMITH_) "$(INTDIR)"\
 "$(INTDIR)\Smith.tlb"
   $(RSC) /l 0x409 /fo"$(INTDIR)/Smith.res" /i "Release" /d "NDEBUG" /d\
 "_AFXDLL" $(SOURCE)


!ELSEIF  "$(CFG)" == "Smith - Win32 Debug"

DEP_RSC_SMITH_=\
	".\Smith.ico"\
	".\Smithctl.bmp"\
	".\Smith.rc2"\
	".\Debug\smith.tlb"\
	

"$(INTDIR)\Smith.res" : $(SOURCE) $(DEP_RSC_SMITH_) "$(INTDIR)"\
 "$(INTDIR)\Smith.tlb"
   $(RSC) /l 0x409 /fo"$(INTDIR)/Smith.res" /i "Debug" /d "_DEBUG" /d "_AFXDLL"\
 $(SOURCE)


!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Debug"

DEP_RSC_SMITH_=\
	".\UniDebug\Smith.tlb"\
	".\Smith.ico"\
	".\Smithctl.bmp"\
	".\Resource.h"\
	".\Smith.rc2"\
	

"$(INTDIR)\Smith.res" : $(SOURCE) $(DEP_RSC_SMITH_) "$(INTDIR)"\
 "$(INTDIR)\Smith.tlb"
   $(RSC) /l 0x409 /fo"$(INTDIR)/Smith.res" /i "UniDebug" /d "_DEBUG" /d\
 "_AFXDLL" $(SOURCE)


!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Release"

DEP_RSC_SMITH_=\
	".\Smith.ico"\
	".\Smithctl.bmp"\
	".\Resource.h"\
	".\Smith.rc2"\
	".\UniRelease\Smith.tlb"\
	

"$(INTDIR)\Smith.res" : $(SOURCE) $(DEP_RSC_SMITH_) "$(INTDIR)"\
 "$(INTDIR)\Smith.tlb"
   $(RSC) /l 0x409 /fo"$(INTDIR)/Smith.res" /i "UniRelease" /d "NDEBUG" /d\
 "_AFXDLL" $(SOURCE)


!ENDIF 

# End Source File
################################################################################
# Begin Source File

SOURCE=.\Smith.odl

!IF  "$(CFG)" == "Smith - Win32 Release"


"$(OUTDIR)\Smith.tlb" : $(SOURCE) "$(OUTDIR)"
   $(MTL) /nologo /D "NDEBUG" /tlb "$(OUTDIR)/Smith.tlb" /win32 $(SOURCE)


!ELSEIF  "$(CFG)" == "Smith - Win32 Debug"


"$(OUTDIR)\Smith.tlb" : $(SOURCE) "$(OUTDIR)"
   $(MTL) /nologo /D "_DEBUG" /tlb "$(OUTDIR)/Smith.tlb" /win32 $(SOURCE)


!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Debug"


"$(OUTDIR)\Smith.tlb" : $(SOURCE) "$(OUTDIR)"
   $(MTL) /nologo /D "_DEBUG" /tlb "$(OUTDIR)/Smith.tlb" /win32 $(SOURCE)


!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Release"


"$(OUTDIR)\Smith.tlb" : $(SOURCE) "$(OUTDIR)"
   $(MTL) /nologo /D "NDEBUG" /tlb "$(OUTDIR)/Smith.tlb" /win32 $(SOURCE)


!ENDIF 

# End Source File
################################################################################
# Begin Source File

SOURCE=.\Smithctl.cpp

!IF  "$(CFG)" == "Smith - Win32 Release"

DEP_CPP_SMITHC=\
	".\Smith.h"\
	".\smithctl.h"\
	".\Smithppg.h"\
	".\stdafx.h"\
	

"$(INTDIR)\Smithctl.obj" : $(SOURCE) $(DEP_CPP_SMITHC) "$(INTDIR)"\
 "$(INTDIR)\smith.pch"


!ELSEIF  "$(CFG)" == "Smith - Win32 Debug"

DEP_CPP_SMITHC=\
	".\stdafx.h"\
	".\Smith.h"\
	".\smithctl.h"\
	".\Smithppg.h"\
	

"$(INTDIR)\Smithctl.obj" : $(SOURCE) $(DEP_CPP_SMITHC) "$(INTDIR)"\
 "$(INTDIR)\smith.pch"


!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Debug"

DEP_CPP_SMITHC=\
	".\Smith.h"\
	".\smithctl.h"\
	".\Smithppg.h"\
	".\stdafx.h"\
	

"$(INTDIR)\Smithctl.obj" : $(SOURCE) $(DEP_CPP_SMITHC) "$(INTDIR)"\
 "$(INTDIR)\smith.pch"


!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Release"

DEP_CPP_SMITHC=\
	".\Smith.h"\
	".\smithctl.h"\
	".\Smithppg.h"\
	".\stdafx.h"\
	

"$(INTDIR)\Smithctl.obj" : $(SOURCE) $(DEP_CPP_SMITHC) "$(INTDIR)"\
 "$(INTDIR)\smith.pch"


!ENDIF 

# End Source File
################################################################################
# Begin Source File

SOURCE=.\Smithppg.cpp

!IF  "$(CFG)" == "Smith - Win32 Release"

DEP_CPP_SMITHP=\
	".\Smith.h"\
	".\Smithppg.h"\
	".\stdafx.h"\
	

"$(INTDIR)\Smithppg.obj" : $(SOURCE) $(DEP_CPP_SMITHP) "$(INTDIR)"\
 "$(INTDIR)\smith.pch"


!ELSEIF  "$(CFG)" == "Smith - Win32 Debug"

DEP_CPP_SMITHP=\
	".\stdafx.h"\
	".\Smith.h"\
	".\Smithppg.h"\
	

"$(INTDIR)\Smithppg.obj" : $(SOURCE) $(DEP_CPP_SMITHP) "$(INTDIR)"\
 "$(INTDIR)\smith.pch"


!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Debug"

DEP_CPP_SMITHP=\
	".\Smith.h"\
	".\Smithppg.h"\
	".\stdafx.h"\
	

"$(INTDIR)\Smithppg.obj" : $(SOURCE) $(DEP_CPP_SMITHP) "$(INTDIR)"\
 "$(INTDIR)\smith.pch"


!ELSEIF  "$(CFG)" == "Smith - Win32 Unicode Release"

DEP_CPP_SMITHP=\
	".\Smith.h"\
	".\Smithppg.h"\
	".\stdafx.h"\
	

"$(INTDIR)\Smithppg.obj" : $(SOURCE) $(DEP_CPP_SMITHP) "$(INTDIR)"\
 "$(INTDIR)\smith.pch"


!ENDIF 

# End Source File
# End Target
# End Project
################################################################################
