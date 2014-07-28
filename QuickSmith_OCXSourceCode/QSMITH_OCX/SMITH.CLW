; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=CSmithCtrl
LastTemplate=COlePropertyPage
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "smith.h"
VbxHeaderFile=smith.h
VbxImplementationFile=smith.cpp
CDK=Y

ClassCount=3
Class1=CSmithCtrl
Class2=CSmithPropPage

ResourceCount=3
Resource1=IDD_PROPPAGE_SMITH
LastPage=0
Resource2=IDD_PROPPAGE_DATA
Class3=CDataPropPage
Resource3=IDD_ABOUTBOX_SMITH

[CLS:CSmithCtrl]
Type=0
HeaderFile=smithctl.h
ImplementationFile=smithctl.cpp
Filter=W
LastObject=CSmithCtrl
BaseClass=COleControl
VirtualFilter=WC

[CLS:CSmithPropPage]
Type=0
HeaderFile=smithppg.h
ImplementationFile=smithppg.cpp
Filter=D
LastObject=CSmithPropPage
BaseClass=COlePropertyPage
VirtualFilter=idWC

[DLG:IDD_ABOUTBOX_SMITH]
Type=1
Class=?
ControlCount=4
Control1=IDC_STATIC,static,1342308352
Control2=IDC_STATIC,static,1342308352
Control3=IDOK,button,1342373889
Control4=IDC_STATIC,static,1342308352

[DLG:IDD_PROPPAGE_SMITH]
Type=1
Class=CSmithPropPage
ControlCount=14
Control1=IDC_SHOWADMITTANCE,button,1342242819
Control2=IDC_STATIC,static,1342308352
Control3=IDC_VSWRCIRCLE,edit,1350631552
Control4=IDC_STATIC,static,1342308352
Control5=IDC_QCIRCLE,edit,1350631552
Control6=IDC_SHOWMARKER,button,1342242819
Control7=IDC_STATIC,static,1342308352
Control8=IDC_MARKERM,edit,1350631552
Control9=IDC_STATIC,static,1342308352
Control10=IDC_MARKERQ,edit,1350631552
Control11=IDC_STATIC,static,1342308352
Control12=IDC_Z0,edit,1350631552
Control13=IDCRESETALL,button,1342242819
Control14=IDCREDRAW,button,1342242819

[DLG:IDD_PROPPAGE_DATA]
Type=1
Class=CDataPropPage
ControlCount=24
Control1=IDC_STATIC,static,1342308352
Control2=IDC_CIRCLEINDEX,edit,1350631552
Control3=IDC_STATIC,static,1342308352
Control4=IDC_CIRCLEM,edit,1350631552
Control5=IDC_STATIC,static,1342308352
Control6=IDC_CIRCLEQ,edit,1350631552
Control7=IDC_STATIC,static,1342308352
Control8=IDC_CIRCLER,edit,1350631552
Control9=IDC_CIRCLE,button,1342177287
Control10=IDC_STATIC,static,1342308352
Control11=IDC_DATAM,edit,1350631552
Control12=IDC_STATIC,static,1342308352
Control13=IDC_DATAQ,edit,1350631552
Control14=IDC_STATIC,static,1342308352
Control15=IDC_NUMPOINTS,edit,1350631552
Control16=IDC_STATIC,static,1342308352
Control17=IDC_NUMSETS,edit,1350631552
Control18=IDC_STATIC,button,1342177287
Control19=IDC_STATIC,static,1342308352
Control20=IDC_THISPOINT,edit,1350631552
Control21=IDC_STATIC,static,1342308352
Control22=IDC_THISSET,edit,1350631552
Control23=IDC_SWEEP,button,1342242851
Control24=IDC_DATARESET,button,1342242851

[CLS:CDataPropPage]
Type=0
HeaderFile=Datappg.h
ImplementationFile=Datappg.cpp
BaseClass=COlePropertyPage
Filter=D
LastObject=CDataPropPage
VirtualFilter=idWC

