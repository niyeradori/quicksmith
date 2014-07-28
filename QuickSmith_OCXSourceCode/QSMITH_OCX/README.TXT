========================================================================
		OLE Custom Control DLL : SMITH
========================================================================

ControlWizard has created this project for your SMITH OLE Custom
Control DLL, which contains 1 Control.

This skeleton project not only demonstrates the basics of writing an OLE
Custom Control, but is also a starting point for writing the specific
features of your control.

This file contains a summary of what you will find in each of the files
that make up your SMITH OLE Custom Control DLL.


SMITH.MAK
	The project makefile for building your 16-bit OLE Custom Control.

SMITH32.MAK
	The project makefile for building your 32-bit OLE Custom Control.

MAKEFILE
	A makefile that makes it easy to run NMAKE from the command prompt.
	Use the following parameters with NMAKE:
		DEBUG=0     Builds retail version
		DEBUG=1     Builds debug version (default)
		WIN32=0     Builds 16-bit version (default)
		WIN32=1     Builds 32-bit version
		UNICODE=0   Builds ANSI/DBCS version (default for WIN32=0)
		UNICODE=1   Build Unicode version (default for WIN32=1)

SMITH.H
	This is the main include file for the OLE Custom Control DLL.  It
	includes other project-specific includes such as RESOURCE.H.

SMITH.CPP
	This is the main source file that contains the OLE DLL initialization,
	termination and other bookkeeping.

SMITH.RC
	This is a listing of the Microsoft Windows resources that the project
	uses.  This file can be directly edited with the Visual C++ resource
	editor.

SMITH.RC2
	This file contains resources that are not edited by the resource editor.
	Initially this contains a VERSIONINFO resource that you can customize
	for your OLE Custom Control DLL, and a TYPELIB resource for your DLL's
	type library.  You should place other manually maintained resources in
	this file.

SMITH.DEF
	This file contains information about the OLE Custom Control DLL that
	must be provided to run with Microsoft Windows.  It defines parameters
	such as the name and description of the DLL, and the size of the initial
	local heap.  The numbers in this file are typical for OLE Custom Control
	DLLs.

SMITH32.DEF
	This is a version of SMITH.DEF for when building a 32-bit version of
	the OLE Custom Control DLL.

SMITH.CLW
	This file contains information used by ClassWizard to edit existing
	classes or add new classes.  ClassWizard also uses this file to store
	information needed to generate and edit message maps and dialog data
	maps and to generate prototype member functions.

SMITH.ODL
	This file contains the Object Description Language source code for the
	type library of your OLE Control.

SMITH.ICO
	This file contains an icon that will appear in the About box.  This icon
	is included by the main resource file SMITH.RC.

/////////////////////////////////////////////////////////////////////////////

SMITHCTL.H
	This file contains the declaration of the CSmithCtrl C++ class.

SMITHCTL.CPP
	This file contains the implementation of the CSmithCtrl C++ class.

SMITHPPG.H
	This file contains the declaration of the CSmithPropPage C++ class.

SMITHPPG.CPP
	This file contains the implementation of the CSmithPropPage C++ class.

SMITHCTL.BMP
	This file contains a bitmap that a container will use to represent the
	CSmithCtrl control when it appears on a tool palette.  This bitmap
	is included by the main resource file SMITH.RC.

/////////////////////////////////////////////////////////////////////////////
Other standard files:

STDAFX.H, STDAFX.CPP
	These files are used to build a precompiled header (PCH) file
	named STDAFX.PCH and a precompiled types (PCT) file named STDAFX.OBJ.

RESOURCE.H
	This is the standard header file, which defines new resource IDs.
	App Studio reads and updates this file.

/////////////////////////////////////////////////////////////////////////////
Other notes:

ControlWizard uses "TODO:" to indicate parts of the source code you
should add to or customize.

/////////////////////////////////////////////////////////////////////////////
