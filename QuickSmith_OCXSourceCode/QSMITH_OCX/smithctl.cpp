/*
    QuickSmith - Smith Chart Program
    Copyright (C) <2013>  <Nathan Iyer>

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.

*/

// smithctl.cpp : Implementation of the CSmithCtrl OLE control class.

#include "stdafx.h"
#include "smith.h"
#include "smithctl.h"
#include "smithppg.h"
#include "datappg.h"
#include "math.h"
#include "comcat.h" // used for internet aware registeration
#include "objsafe.h" //used for internet aware registeration


#ifdef _DEBUG
#undef THIS_FILE
static char BASED_CODE THIS_FILE[] = __FILE__;
#endif

// Define statements are in header file "smithctl.h"



IMPLEMENT_DYNCREATE(CSmithCtrl, COleControl)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CSmithCtrl, COleControl)
	//{{AFX_MSG_MAP(CSmithCtrl)
	ON_WM_LBUTTONDBLCLK()
	ON_WM_MOUSEMOVE()
	//}}AFX_MSG_MAP
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Dispatch map

BEGIN_DISPATCH_MAP(CSmithCtrl, COleControl)
	//{{AFX_DISPATCH_MAP(CSmithCtrl)
	DISP_PROPERTY_EX(CSmithCtrl, "ShowHistory", GetShowHistory, SetShowHistory, VT_BOOL)
	DISP_PROPERTY_EX(CSmithCtrl, "ShowAdmittance", GetShowAdmittance, SetShowAdmittance, VT_BOOL)
	DISP_PROPERTY_EX(CSmithCtrl, "NumPoints", GetNumPoints, SetNumPoints, VT_I2)
	DISP_PROPERTY_EX(CSmithCtrl, "DataM", GetDataM, SetDataM, VT_R4)
	DISP_PROPERTY_EX(CSmithCtrl, "DataQ", GetDataQ, SetDataQ, VT_R4)
	DISP_PROPERTY_EX(CSmithCtrl, "ThisPoint", GetThisPoint, SetThisPoint, VT_I2)
	DISP_PROPERTY_EX(CSmithCtrl, "DataReset", GetDataReset, SetDataReset, VT_BOOL)
	DISP_PROPERTY_EX(CSmithCtrl, "VSWRCircle", GetVSWRCircle, SetVSWRCircle, VT_R4)
	DISP_PROPERTY_EX(CSmithCtrl, "QCircle", GetQCircle, SetQCircle, VT_R4)
	DISP_PROPERTY_EX(CSmithCtrl, "ShowMarker", GetShowMarker, SetShowMarker, VT_BOOL)
	DISP_PROPERTY_EX(CSmithCtrl, "ResetAll", GetResetAll, SetResetAll, VT_BOOL)
	DISP_PROPERTY_EX(CSmithCtrl, "MarkerM", GetMarkerM, SetMarkerM, VT_R4)
	DISP_PROPERTY_EX(CSmithCtrl, "MarkerQ", GetMarkerQ, SetMarkerQ, VT_R4)
	DISP_PROPERTY_EX(CSmithCtrl, "Redraw", GetRedraw, SetRedraw, VT_BOOL)
	DISP_PROPERTY_EX(CSmithCtrl, "Sweep", GetSweep, SetSweep, VT_BOOL)
	DISP_PROPERTY_EX(CSmithCtrl, "Z0", GetZ0, SetZ0, VT_R4)
	DISP_PROPERTY_EX(CSmithCtrl, "CircleM", GetCircleM, SetCircleM, VT_R4)
	DISP_PROPERTY_EX(CSmithCtrl, "CircleQ", GetCircleQ, SetCircleQ, VT_R4)
	DISP_PROPERTY_EX(CSmithCtrl, "CircleR", GetCircleR, SetCircleR, VT_R4)
	DISP_PROPERTY_EX(CSmithCtrl, "CircleColor", GetCircleColor, SetCircleColor, VT_COLOR)
	DISP_PROPERTY_EX(CSmithCtrl, "CircleIndex", GetCircleIndex, SetCircleIndex, VT_I2)
	DISP_PROPERTY_EX(CSmithCtrl, "ThisSet", GetThisSet, SetThisSet, VT_I2)
	DISP_PROPERTY_EX(CSmithCtrl, "NumSets", GetNumSets, SetNumSets, VT_I2)
	DISP_DEFVALUE(CSmithCtrl, "ShowAdmittance")
	DISP_STOCKPROP_BACKCOLOR()
	DISP_STOCKPROP_FONT()
	DISP_STOCKPROP_HWND()
	DISP_PROPERTY_EX_ID(CSmithCtrl, "ForeColor", DISPID_FORECOLOR, GetForeColor, SetForeColor, VT_COLOR)
	//}}AFX_DISPATCH_MAP
	DISP_FUNCTION_ID(CSmithCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()


/////////////////////////////////////////////////////////////////////////////
// Event map

BEGIN_EVENT_MAP(CSmithCtrl, COleControl)
	//{{AFX_EVENT_MAP(CSmithCtrl)
	// NOTE - ClassWizard will add and remove event map entries
	//    DO NOT EDIT what you see in these blocks of generated code !
	EVENT_CUSTOM("ClickIn", FireClickIn,  VTS_PR4 VTS_PR4)
	EVENT_CUSTOM("ClickOut", FireClickOut,  VTS_NONE)
	EVENT_CUSTOM("MouseMove", FireMouseMove,  VTS_PR4 VTS_PR4)
	EVENT_STOCK_KEYDOWN()
	EVENT_STOCK_KEYPRESS()
	EVENT_STOCK_KEYUP()

	//}}AFX_EVENT_MAP
END_EVENT_MAP()


/////////////////////////////////////////////////////////////////////////////
// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CSmithCtrl, 4)
	PROPPAGEID(CSmithPropPage::guid)
	PROPPAGEID(CDataPropPage::guid)
	PROPPAGEID(CLSID_CColorPropPage)
	PROPPAGEID(CLSID_CFontPropPage)
END_PROPPAGEIDS(CSmithCtrl)


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CSmithCtrl, "SMITH.SmithCtrl.1",
	0x5ae352a0, 0xe9a4, 0x11d0, 0x93, 0x6d, 0x44, 0x45, 0x53, 0x54, 0x0, 0x0)


/////////////////////////////////////////////////////////////////////////////
// Type library ID and version

IMPLEMENT_OLETYPELIB(CSmithCtrl, _tlid, _wVerMajor, _wVerMinor)


/////////////////////////////////////////////////////////////////////////////
// Interface IDs

const IID BASED_CODE IID_DSmith =
		{ 0x5ae352a1, 0xe9a4, 0x11d0, { 0x93, 0x6d, 0x44, 0x45, 0x53, 0x54, 0x0, 0x0 } };
const IID BASED_CODE IID_DSmithEvents =
		{ 0x5ae352a2, 0xe9a4, 0x11d0, { 0x93, 0x6d, 0x44, 0x45, 0x53, 0x54, 0x0, 0x0 } };


/////////////////////////////////////////////////////////////////////////////
// Control type information

static const DWORD BASED_CODE _dwSmithOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CSmithCtrl, IDS_SMITH, _dwSmithOleMisc)


/////////////////////////////////////////////////////////////////////////////
// CSmithCtrl::CSmithCtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CSmithCtrl

// These three functions below are for internet aware controls
HRESULT CreateComponentCategory( CATID catid, WCHAR* catDescription )
{
   ICatRegister* pcr = NULL ;
   HRESULT hr = S_OK ;

   // Create an instance of the category manager.
   hr = CoCreateInstance( CLSID_StdComponentCategoriesMgr,
                          NULL,
                          CLSCTX_INPROC_SERVER,
                          IID_ICatRegister,
                          (void**)&pcr );
   if (FAILED(hr))
      return hr;

   CATEGORYINFO catinfo;
   catinfo.catid = catid;
   // English locale ID in hex
   catinfo.lcid = 0x0409;

   int len = wcslen(catDescription);
   wcsncpy( catinfo.szDescription, catDescription, len );
   catinfo.szDescription[len] = '\0';

   hr = pcr->RegisterCategories( 1, &catinfo );
   pcr->Release();

   return hr;
}

HRESULT RegisterCLSIDInCategory( REFCLSID clsid, CATID catid )
{
   ICatRegister* pcr = NULL ;
   HRESULT hr = S_OK ;

   // Create an instance of the category manager.
   hr = CoCreateInstance( CLSID_StdComponentCategoriesMgr,
                          NULL,
                          CLSCTX_INPROC_SERVER,
                          IID_ICatRegister,
                          (void**)&pcr );
   if (SUCCEEDED(hr))
   {
      CATID rgcatid[1];
      rgcatid[0] = catid;
      hr = pcr->RegisterClassImplCategories( clsid, 1, rgcatid );
   }

   if ( pcr != NULL )
      pcr->Release();

   return hr;
}


HRESULT UnregisterCLSIDInCategory( REFCLSID clsid, CATID catid )
{
   ICatRegister* pcr = NULL ;
   HRESULT hr = S_OK ;

   // Create an instance of the category manager.
   hr = CoCreateInstance( CLSID_StdComponentCategoriesMgr,
                          NULL,
                          CLSCTX_INPROC_SERVER,
                          IID_ICatRegister,
                          (void**)&pcr );
   if (SUCCEEDED(hr))
   {
      CATID rgcatid[1];
      rgcatid[0] = catid;
      hr = pcr->UnRegisterClassImplCategories( clsid, 1, rgcatid );
   }

   if ( pcr != NULL )
      pcr->Release();

   return hr;
}



BOOL CSmithCtrl::CSmithCtrlFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
	{
	  CreateComponentCategory( CATID_Control,
                               L"Controls" );
      RegisterCLSIDInCategory( m_clsid,
                               CATID_Control );

      CreateComponentCategory( CATID_SafeForInitializing,
                               L"Controls safely initializable from persistent data" );
      RegisterCLSIDInCategory( m_clsid,
                               CATID_SafeForInitializing );

      CreateComponentCategory( CATID_SafeForScripting,
                               L"Controls that are safely scriptable" );
      RegisterCLSIDInCategory( m_clsid,
                               CATID_SafeForScripting );

      CreateComponentCategory( CATID_PersistsToPropertyBag,
                               L"Support initialize via PersistPropertyBag" );
      RegisterCLSIDInCategory( m_clsid,
                               CATID_PersistsToPropertyBag );
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(), // Commented to indicate matching registry entry
			m_clsid,				//CLSID
			m_lpszProgID,			//ProgID
			IDS_SMITH,				// Textual Control Name
			IDB_SMITH,				// ToolboxBitmap32
			FALSE,                  //  Not insertable
			_dwSmithOleMisc,		// Misc Status
			_tlid,					// TypeLib
			_wVerMajor,
			_wVerMinor);			// Combined to create version
    }
	else
    {
	  UnregisterCLSIDInCategory( m_clsid,
                                 CATID_Control );
      UnregisterCLSIDInCategory( m_clsid,
                                 CATID_SafeForInitializing );
      UnregisterCLSIDInCategory( m_clsid,
                                 CATID_SafeForScripting );
      UnregisterCLSIDInCategory( m_clsid,
                                 CATID_PersistsToPropertyBag );
	  return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
	}
	}

/////////////////////////////////////////////////////////////////////////////
// Licensing strings

static const TCHAR BASED_CODE _szLicFileName[] = _T("smith.lic");

static const WCHAR BASED_CODE _szLicString[] = L"V.Iyer (iyer@emi.net), Copyright (c) 1997 ";


/////////////////////////////////////////////////////////////////////////////
// CSmithCtrl::CSmithCtrlFactory::VerifyUserLicense -
// Checks for existence of a user license

BOOL CSmithCtrl::CSmithCtrlFactory::VerifyUserLicense()
{
	return AfxVerifyLicFile(AfxGetInstanceHandle(), _szLicFileName,
		_szLicString);
}


/////////////////////////////////////////////////////////////////////////////
// CSmithCtrl::CSmithCtrlFactory::GetLicenseKey -
// Returns a runtime licensing key

BOOL CSmithCtrl::CSmithCtrlFactory::GetLicenseKey(DWORD dwReserved,
	BSTR FAR* pbstrKey)
{ DWORD a;
  a = dwReserved;
	if (pbstrKey == NULL)
		return FALSE;

	*pbstrKey = SysAllocString(_szLicString);
	return (*pbstrKey != NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CSmithCtrl::CSmithCtrl - Constructor

CSmithCtrl::CSmithCtrl()
{
	InitializeIIDs(&IID_DSmith, &IID_DSmithEvents);
	OnResetState();

     mHBMImage  = NULL;
	 mHBMMask   = NULL;
	 first_draw =1;
	// TODO: Initialize your control's instance data here.
}


/////////////////////////////////////////////////////////////////////////////
// CSmithCtrl::~CSmithCtrl - Destructor

CSmithCtrl::~CSmithCtrl()
{
	// TODO: Cleanup your control's instance data here.
	delete mHBMImage;
	delete mHBMMask;
	GlobalFree((GLOBALHANDLE)hGmemM);
    GlobalFree((GLOBALHANDLE)hGmemQ);
}


/////////////////////////////////////////////////////////////////////////////
// CSmithCtrl::OnDraw - Drawing function

void CSmithCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{


   CDC dcMem,dcMem1,dcMem2;
   CBitmap bitOff;
   CRect rcBoundsDP(rcBounds);
   CRect rcInvalidDP(rcInvalid);
   CBrush* pOldBr;

   CBrush bkBrush(TranslateColor(GetBackColor()));

	if(first_draw)   // create sprite bitmaps
	{
    int Row, Col;
    pdc->SetMapMode(MM_TEXT) ;
	pdc->SetWindowOrg(0,0) ;
    pdc->SetViewportOrg(0,0) ;
    dcMem1.CreateCompatibleDC(pdc);
	mHBMImage = new CBitmap;
	mHBMImage->CreateCompatibleBitmap(pdc,BWIDTH,BHEIGTH);

	dcMem1.SelectObject(mHBMImage);
    dcMem1.SelectObject(GetStockObject(BLACK_PEN));
    dcMem1.SelectObject(GetStockObject(BLACK_BRUSH));
    dcMem1.Rectangle(0,0,BWIDTH,BHEIGTH);
    dcMem1.SelectObject(GetStockObject(WHITE_PEN));
    dcMem1.SelectObject(GetStockObject(WHITE_BRUSH));
	dcMem1.Ellipse(0,0,BWIDTH,BHEIGTH);
    //Make a mask bitmap
	dcMem2.CreateCompatibleDC(pdc);
	mHBMMask = new CBitmap;
	mHBMMask->CreateCompatibleBitmap(pdc,BWIDTH,BHEIGTH);

	dcMem2.SelectObject(mHBMMask);
	for (Row = 0; Row < BHEIGTH; ++Row)
      for (Col = 0; Col < BWIDTH; ++Col)
         if (dcMem1.GetPixel (Col, Row) == RGB (0,0,0)) // black
            dcMem2.SetPixel (Col, Row, RGB (255,255,255));
         else
            dcMem2.SetPixel (Col, Row, RGB (255,255,0));   //not black

	DeleteDC(dcMem1);
	DeleteDC(dcMem2);
	SpotDataM=GetDataM();    // initialize spot data for serialization consistency
    SpotDataQ=GetDataQ();   // VB remembers the saved state and OCX dont.
    first_draw=0;           //  To Do -include the sweep mode data too
	}
   // Convert bounds to device units.

   pdc->LPtoDP(&rcBoundsDP) ;
   pdc->LPtoDP(&rcInvalidDP);
   // The bitmap bounds have 0,0 in the upper-left corner.
   CRect rcBitmapBounds(0,0,rcBoundsDP.Width(),rcBoundsDP.Height()) ;

   // Create a DC that is compatible with the screen.
   dcMem.CreateCompatibleDC(pdc) ;

   // Create a really compatible bitmap.
   bitOff.CreateBitmap(rcBitmapBounds.Width(),
                       rcBitmapBounds.Height(),
                       pdc->GetDeviceCaps(PLANES),
                       pdc->GetDeviceCaps(BITSPIXEL),
					   NULL);


   // Select the bitmap into the memory DC.
   CBitmap* pOldBitmap = dcMem.SelectObject(&bitOff) ;

   // Save the memory DC state, since DrawMe might change it.
   int iSavedDC = dcMem.SaveDC();


   // Draw our control on the memory DC.
    pOldBr= dcMem.SelectObject(&bkBrush);
    dcMem.FillRect(rcBitmapBounds,&bkBrush);  // fill the area with BackColor between the Smith circle  and window

	MapSmith(&dcMem, rcBitmapBounds);  //  Sets client coordinates
    DrawSmith(&dcMem, rcBitmapBounds);


   // Restore the DC, since DrawMe might have changed mapping modes.
   dcMem.RestoreDC(iSavedDC) ;

   // We don't know what mapping mode pdc is using.
   // BitBlt uses logical coordinates.
   // Easiest thing is to change to MM_TEXT.
   pdc->SetMapMode(MM_TEXT) ;
   pdc->SetWindowOrg(0,0) ;
   pdc->SetViewportOrg(0,0) ;

   // Blt the memory device context to the screen. use rcInvalid to improve performance

   pdc->BitBlt( rcInvalidDP.left,
                rcInvalidDP.top,
                rcInvalidDP.Width(),
                rcInvalidDP.Height(),
                &dcMem,
                rcInvalidDP.left-rcBoundsDP.left,
                rcInvalidDP.top -rcBoundsDP.top,
                SRCCOPY) ;

   // Clean up.
   dcMem.SelectObject(pOldBitmap) ;
   dcMem.SelectObject(pOldBr);

}

/* Metafile Implementation */
void CSmithCtrl::OnDrawMetafile(CDC* pdc, const CRect& rcBounds)
{
POINT pt;
ASSERT (rcBounds.left == 0);ASSERT (rcBounds.top == 0);
CSize size = rcBounds.Size();
int iShortSide = min(size.cx,size.cy);
CPoint ptCenter(rcBounds.left+size.cx/2, rcBounds.top+size.cy/2);
double dbConvert = 2*AXIS_RANGE/(double) iShortSide;
pdc->SetWindowOrg(-int(ptCenter.x*dbConvert+10), int(ptCenter.y*dbConvert+40))  ; // Set logical range for each axis
pdc->SetWindowExt(int(size.cx*dbConvert+20),-int(size.cy*dbConvert+80))  ; // Set logical range for each axis/
pdc->SetViewportExt((rcBounds.right-rcBounds.left)/2,(rcBounds.top-rcBounds.bottom)/2);
pdc->SetViewportOrg((rcBounds.left+rcBounds.right)/2,(rcBounds.top+rcBounds.bottom)/2);

CBitmap*mHBMTemp = mHBMImage;	/* disable sprite drawing in smithdraw function */
mHBMImage = NULL;				/* metafile does not accept change of Mapping modes hence a seperate routine */
DrawSmith( pdc, rcBounds);      /* which draws the sprite without changing the mapping mode */
	{
    CBrush YellowBrush;
    CPen YellowPen;
	pt.x =  (short) (SpotDataM * AXIS_RANGE * cos(Pi * SpotDataQ/180));
    pt.y = (short) (SpotDataM * AXIS_RANGE * sin(Pi * SpotDataQ/180));
	YellowPen.CreatePen(PS_SOLID,1,RGB(255,255,0));
	YellowBrush.CreateSolidBrush(RGB(255,255,0));
	CBrush* OldBrush = (CBrush*)pdc->SelectStockObject(BLACK_PEN);
	CPen* OldPen= (CPen*)pdc->SelectObject(&YellowBrush);
	pdc->SetROP2(R2_XORPEN);
	pdc->Circle(pt.x,pt.y,20);
	pdc->SelectObject (OldPen);		// first select a pen in device context
	pdc->SelectObject (OldBrush);
	pdc->SetROP2(R2_COPYPEN);
	YellowPen.DeleteObject();
	YellowBrush.DeleteObject();
	}
mHBMImage = mHBMTemp ;			/* enable back sprite draw in regular function */

}

void CSmithCtrl::MapSmith(CDC* pdc, CRect rc)  //  Sets client coordinates
{
	// unlike VBX do not assume left and top of client window to be zero
	pdc->SetMapMode(MM_ISOTROPIC) ;
    pdc->SetWindowExt(AXIS_RANGE+20,AXIS_RANGE+80)  ; // Set logical range for each axis
    pdc->SetViewportExt((rc.right-rc.left)/2,(rc.top-rc.bottom)/2);//logical extends 2*1000,2*1000 is the the device
              									 //extends(cxclient,cyclient) and decrease as we go down
    pdc->SetViewportOrg((rc.left+rc.right)/2,(rc.top+rc.bottom)/2); // logical center(0,0) to be the client center
}
void CSmithCtrl::DrawSmith(CDC* pdc,CRect rcBounds)
{
	short r = AXIS_RANGE;         /* radius set to AXIS_RANGE */
	CPen* pOldPen;
	CPen RedPen,GreenPen,GreyDotPen,BluePen,PinkPen;
	CDC dcMem;
	POINT pt;
    int i;


	pdc->SelectStockObject(WHITE_BRUSH);
	RedPen.CreatePen(PS_SOLID,1,RGB(255,0,0));
	GreenPen.CreatePen(PS_SOLID,1,RGB(0,255,0));
	pOldPen =(CPen*)pdc->SelectObject(&RedPen);
  	pdc->Circle(0,0,r);
    pdc->SetBkMode(TRANSPARENT);
    pdc->Circle(r/3,0,2*r/3);			// Rn = 0.5
    pdc->Circle(r/2,0,r/2);				// Rn = 1
    pdc->Circle(2*r/3,0,r/3);			// Rn = 2
    pdc->SelectObject(&GreenPen);
    pdc->MyArc(r,r,r,180,270);			// Xn = 1
    pdc->MyArc(r,-r,r,90,180);
    pdc->MyArc(r,r/2,r/2,143,270);		// Xn = 2
    pdc->MyArc(r,-r/2,r/2,90,217);
    pdc->MyArc(r,2*r,2*r,217,270);		// Xn = 0.5
    pdc->MyArc(r,-2*r,2*r,90,143);
    pdc->MoveTo(-r,0) ;					// Center Line
    pdc->LineTo(r,0);
	pdc->SetTextColor(RGB(0,0,0));//Black
    pdc->SetBkMode(TRANSPARENT);

	CFont* pOldFont = SelectStockFont(pdc);
	if(!mHBMImage) // Only execute for metafile draw
	{
	pdc->SetTextAlign(TA_CENTER | TA_TOP);
	pdc->TextOut(0,0," r=1",4);
    pdc->TextOut(r/3,0," 2",2);
    pdc->TextOut(-r/3,0," 0.5",4);
    pdc->SetTextAlign(TA_BOTTOM);pdc->TextOut(0,r,"x=1",3);
    pdc->SetTextAlign(TA_TOP);pdc->TextOut(0,5-r,"x=-1",4);
    pdc->SetTextAlign(TA_BOTTOM|TA_LEFT);pdc->TextOut((short)(r*cos(Pi*53/180)),(short)(r*sin(Pi*53/180))," 2",2);
    pdc->SetTextAlign(TA_LEFT);pdc->TextOut((short)(r*cos(Pi*307/180)),(short)(r*sin(Pi*307/180)),"-2",2);
    pdc->SetTextAlign(TA_BOTTOM|TA_RIGHT);pdc->TextOut((short)(r*cos(Pi*127/180)),(short)(r*sin(Pi*127/180))," 0.5",4);
    pdc->SetTextAlign(TA_RIGHT|TA_TOP);pdc->TextOut((short)(r*cos(Pi*233/180)),(short)(r*sin(Pi*233/180)),"-0.5",4);
    }


	if (m_bShowAdmittance)
       {
		 GreyDotPen.CreatePen(PS_DOT,1,RGB(128,128,128));
         pdc->SelectObject (&GreyDotPen);	// select grey dotted Pen
         //Constant Conductance cicles
         pdc->MyArc(-2*r/3,0,r/3,0,360);	//  Cn = 0.5
         pdc->MyArc(-r/2,0,r/2,0,360);		// Cn = 1
         pdc->MyArc(-r/3,0,2*r/3,0,360);	// Cn = 2
         //constant Suseptance circles
         pdc->MyArc(-r,r/2,r/2,-90,37);		// Bn = 2
         pdc->MyArc(-r,-r/2,r/2,-37,90);    // Bn = -2
		 pdc->MyArc(-r,r,r,-90,0);			// Bn = -1
         pdc->MyArc(-r,-r,r,0,90);			// Bn = 1
		 pdc->MyArc(-r,2*r,2*r,-90,-37);	// Bn = 0.5
         pdc->MyArc(-r,-2*r,2*r,37,90);		// Bn = -0.5
		 pdc->SelectObject (pOldPen);		// first select a pen in device context
		 GreyDotPen.DeleteObject();
	   }

	// draw historical VSWR circles if any
	if((m_bShowHistory == TRUE) && (vswrindex > 0)) // if at least one is stored in index > 0 
	{
	   for (i=1;i<=HISTORYMAX+1;i++)
	   { if (vswrstor[i] >0 ) 
	   { float swr,r;
         short radius;
		 BluePen.CreatePen(PS_DOT,1,RGB(0,0,255));
         swr = vswrstor[i];
         r = ((swr -1)/(1+swr));
         radius = (short)(r*AXIS_RANGE);
         pdc->SelectObject(&BluePen);
         pdc->MyArc(0,0,radius,0,360);
		 pdc->SelectObject (pOldPen);		// first select a pen in device context
		 BluePen.DeleteObject();
	   }
	   }
      }

    // draw current VSWR
    if (m_fpVSWRCircle)
       {
        float swr,r;
        short radius;
		BluePen.CreatePen(PS_SOLID,1,RGB(0,0,255));
        swr = m_fpVSWRCircle;
        r = ((swr -1)/(1+swr));
        radius = (short)(r*AXIS_RANGE);
        pdc->SelectObject(&BluePen);
        pdc->MyArc(0,0,radius,0,360);
		pdc->SelectObject (pOldPen);		// first select a pen in device context
		BluePen.DeleteObject();
       }

	// draw historical Q circles if any
	if(m_bShowHistory && (qindex > 0)) // if at least one is stored in index > 0 
	{
	   for (i=1;i<=HISTORYMAX+1;i++)
	   { if (qstor[i] >0 ) 
	   { float q,r;
         short radius,center ;
	     PinkPen.CreatePen(PS_DOT,1,RGB(255,0,255));
         q =qstor[i];
         r = 1/q;
         r=(float) sqrt(1+r*r);
         radius = (short)abs((short)(r*AXIS_RANGE));
         r=1/q;
         center = (short)abs((short)(r*AXIS_RANGE));
         pdc->SelectObject(&PinkPen);
         pdc->MyQArc2(0,-center,radius);
         pdc->MyQArc1(0,center,radius);
	     pdc->SelectObject (pOldPen);		// first select a pen in device context
	     PinkPen.DeleteObject();
	   }
	   }
      }


    if(m_fpQCircle)
      {
       float q,r;
       short radius,center ;
	   PinkPen.CreatePen(PS_SOLID,1,RGB(255,0,255));
       q =m_fpQCircle;
       r = 1/q;
       r=(float) sqrt(1+r*r);
       radius = (short)abs((short)(r*AXIS_RANGE));
       r=1/q;
       center = (short)abs((short)(r*AXIS_RANGE));
       pdc->SelectObject(&PinkPen);
       pdc->MyQArc2(0,-center,radius);
       pdc->MyQArc1(0,center,radius);
	   pdc->SelectObject (pOldPen);		// first select a pen in device context
	   PinkPen.DeleteObject();
     }

	if(m_fpZ0 && !mHBMImage)  // again a text function perform only for meta draw
	{						// for non meta draw perform in the MMTEXT area
	  int len;
	  char szBuffer[10];
	  pt.x =950;
      pt.y =-1020;
      pdc->SetTextAlign(TA_RIGHT|TA_BOTTOM);
	  len = sprintf(szBuffer,"%.1f",m_fpZ0);
	  pdc->TextOut(pt.x,pt.y,szBuffer,len);
      pdc->SetTextAlign(TA_TOP|TA_LEFT);
	}

	if(m_bShowMarker)
	{
      float r,q;
      int R;
	  r= m_fpMarkerM;
      q= m_fpMarkerQ;
      pt.x=(short) (AXIS_RANGE*r*cos(q*Pi/180));
      pt.y= (short) (AXIS_RANGE*r*sin(q*Pi/180));
	  //pdc->LPtoDP(&pt);          /* here the mapping mode is MM_ISO */
	 // pdc->SetMapMode(MM_TEXT) ; /* MM_TEXT mode draws with consitant size */
     // pdc->SetWindowOrg(0,0) ;
     //  pdc->SetViewportOrg(0,0) ;
      pdc->SelectStockObject(BLACK_PEN);
      //R =7 ;				// Set marker size in device units
	  R=36 ;   //*
      //pdc->Arc(pt.x-R,pt.y+R+1,pt.x+R+1,pt.y-R,pt.x+R,pt.y,pt.x+R,pt.y);
	  pdc->MyArc(pt.x,pt.y,R,0,360); //*
	  pdc->MoveTo(pt.x,pt.y-R);
      pdc->LineTo (pt.x,pt.y+R);
      pdc->MoveTo(pt.x-R,pt.y);
      pdc->LineTo (pt.x+R,pt.y);
	 // MapSmith(pdc, rcBounds);  // restore MM_ISO

   	}

	// draw historical circles if any
	if(m_bShowHistory && (histindex > 0)) // if one is stored histindex > 0 
    {
     short ptx,pty,radius;
     for (i=1;i<=HISTORYMAX+1;i++)
	    {
        radius = (short)(circstor[i].rh * AXIS_RANGE);
      	if(radius > 0)
		     {
          CPen CircleColorPen(PS_DOT,1,circstor[i].cch); // create and select the pen
		      pdc->SelectObject(&CircleColorPen);
           ptx=(short)(AXIS_RANGE * circstor[i].mh * cos( circstor[i].qh*Pi/180));
          pty=(short)(AXIS_RANGE * circstor[i].mh * sin( circstor[i].qh*Pi/180));
          pdc->MyArc(ptx,pty,radius,0,360);
		     }
      }

	  }

	if((short)GetCircleIndex())
    {
     short ptx,pty,radius;
     for (i=1;i<=CIRCLEMAX;i++)
	 {
        radius = (short)(circler[i] * AXIS_RANGE);
      	if(radius > 0)
		{
         CPen CircleColorPen(PS_SOLID,1,circlecolor[i]); // create and select the pen
		 pdc->SelectObject(&CircleColorPen);
         ptx=(short)(AXIS_RANGE * circlem[i] * cos( circleq[i]*Pi/180));
         pty=(short)(AXIS_RANGE * circlem[i] * sin( circleq[i]*Pi/180));
         pdc->MyArc(ptx,pty,radius,0,360);
		}
      }

	}

	if(mHBMImage) /* Throw all MMtext functions here- OnDrawMetafile does not excecute this*/
	{
	int len;
	char szBuffer[10];
	POINT pt1[10];
    pt.x =  (short) (SpotDataM * AXIS_RANGE * cos(Pi * SpotDataQ/180));
    pt.y = (short) (SpotDataM * AXIS_RANGE * sin(Pi * SpotDataQ/180));
	// map the points from iso to text mode first
	pdc->LPtoDP(&pt);          // here the maping mode is MM_ISO
	pt1[1].x = 0;pt1[1].y = 0;pdc->LPtoDP(&pt1[1]);
    pt1[2].x = r/3;pt1[2].y =0;pdc->LPtoDP( &pt1[2]);
    pt1[3].x = -r/3;pt1[3].y = 0;pdc->LPtoDP(&pt1[3]);
    pt1[4].x = 0;pt1[4].y =r;pdc->LPtoDP(&pt1[4]);
    pt1[5].x = 0;pt1[5].y = 5-r;pdc->LPtoDP( &pt1[5]);
    pt1[6].x = (short)(r*cos(Pi*53/180));pt1[6].y = (short)(r*sin(Pi*53/180));pdc->LPtoDP( &pt1[6]);
    pt1[7].x = (short)(r*cos(Pi*307/180));pt1[7].y =(short)(r*sin(Pi*307/180));pdc->LPtoDP(&pt1[7]);
    pt1[8].x = (short)(r*cos(Pi*127/180));pt1[8].y =(short)(r*sin(Pi*127/180));pdc->LPtoDP( &pt1[8]);
    pt1[9].x = (short)(r*cos(Pi*233/180));pt1[9].y =(short)(r*sin(Pi*233/180));pdc->LPtoDP(&pt1[9]);
	pt1[10].x =950; pt1[10].y=-1020; pdc->LPtoDP( &pt1[10]);

	pdc->SetMapMode(MM_TEXT) ; // MM_TEXT mode draws sprirte  with consistent size
    pdc->SetWindowOrg(0,0) ;
    pdc->SetViewportOrg(0,0) ;
    dcMem.CreateCompatibleDC(pdc) ;
    dcMem.SelectObject(mHBMMask) ;  // AND and COPY the Mask bitmap to the screen
    pdc->BitBlt(pt.x-(BWIDTH-1)/2,pt.y-(BHEIGTH-1)/2,BWIDTH,BHEIGTH,&dcMem,0,0,SRCAND);
    dcMem.SelectObject(mHBMImage) ; // INVERT AND COPY the Imagebitmap to the screen
    pdc->BitBlt(pt.x-(BWIDTH-1)/2,pt.y-(BHEIGTH-1)/2,BWIDTH,BHEIGTH,&dcMem,0,0,SRCINVERT);
  	DeleteDC(dcMem);

	/* font size donot register or change in mapmode hence has to be writtenin MMtext mode */
	/* however Metfile draw does not like to see Mapmode commands so disable for Meta draw*/

	pdc->TextOut(pt1[1].x,pt1[1].y," r=1",4);
    pdc->TextOut(pt1[2].x,pt1[2].y," 2",2);
    pdc->TextOut(pt1[3].x,pt1[3].y," 0.5",4);
    pdc->SetTextAlign(TA_BOTTOM);          pdc->TextOut(pt1[4].x,pt1[4].y,"x=1",3);
    pdc->SetTextAlign(TA_TOP);             pdc->TextOut(pt1[5].x,pt1[5].y,"x=-1",4);
    pdc->SetTextAlign(TA_BOTTOM|TA_LEFT);  pdc->TextOut(pt1[6].x,pt1[6].y," 2",2);
    pdc->SetTextAlign(TA_LEFT);            pdc->TextOut(pt1[7].x,pt1[7].y,"-2",2);
    pdc->SetTextAlign(TA_BOTTOM|TA_RIGHT); pdc->TextOut(pt1[8].x,pt1[8].y," 0.5",4);
    pdc->SetTextAlign(TA_RIGHT|TA_TOP);    pdc->TextOut(pt1[9].x,pt1[9].y,"-0.5",4);

	// Print Z0
    if(m_fpZ0)
	{
	pdc->SetTextAlign(TA_RIGHT|TA_BOTTOM);
	len = sprintf(szBuffer,"%.1f",m_fpZ0);
	pdc->TextOut(pt1[10].x,pt1[10].y,szBuffer,len);
    pdc->SetTextAlign(TA_TOP|TA_LEFT);
	}
    MapSmith(pdc, rcBounds);  // restore MM_ISO
	}


   if(m_bSweep)  //Sweep mode
   { int j,k;
	short ptX,ptY,imax;
	CPen PenTraceColor(PS_SOLID,1,TranslateColor(GetForeColor()));

    j=GetNumSets();

	dataM =(float FAR *)GlobalLock(hGmemM);  // get pointer for data
	dataQ= (float FAR *)GlobalLock(hGmemQ);
    for (k=1; k<=j; k++)   // draw untill this set is reached
	{
		imax =(short)thispointmax[k-1]; // imax of previous set set set
		CPen PenTraceColor(PS_SOLID,1,forecolor[k]); // create and select the pen
		pdc->SelectObject(&PenTraceColor);
		ptX =  (short) (*(dataM + ((k-1)*(imax))) *AXIS_RANGE * cos(Pi * *(dataQ+ ((k-1)*(imax)))/180));
		ptY = (short) (*(dataM + ((k-1)*(imax)) ) *AXIS_RANGE * sin(Pi * *(dataQ+((k-1)*(imax)))/180));
   		pdc->MoveTo (ptX,ptY); // from the 0th point then join the rest using lines
   		pdc->SetPixel(ptX,ptY,TranslateColor(GetForeColor()));

		for (i=1; i<thispointmax[k]; i++) // here you want to use Imax of this set
		{
		ptX =  (short) (*(dataM + i + ((k-1)*imax))*AXIS_RANGE * cos(Pi * *(dataQ +i + ((k-1)*imax))/180));
		ptY = (short) (*(dataM + i + ((k-1)*imax))*AXIS_RANGE * sin(Pi * *(dataQ +i + ((k-1)*imax))/180));
		pdc->LineTo (ptX,ptY);
		}
    }
	GlobalUnlock((GLOBALHANDLE)hGmemM);  // release pointer
    GlobalUnlock((GLOBALHANDLE)hGmemQ);

   }
 	pdc->SelectObject(pOldPen);
	pdc->SelectObject(pOldFont);
    RedPen.DeleteObject();     // current pen selected in dc is not RedPen
	GreenPen.DeleteObject();   // or GreenPen
	//TRACE("Draw Smith %x\n",this);
}
/////////////////////////////////////////////////////////////////////////////
// CSmithCtrl::DoPropExchange - Persistence support

void CSmithCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	// TODO: Call PX_ functions for each persistent custom property.


	//if (IsConvertingVBX())
	//	PX_VBXFontConvert(pPX, InternalGetFont());

	PX_Short(pPX, _T("NumPoints"), m_iNumPoints,5);
	PX_Float(pPX, _T("DataM"), m_fpDataM,0);
	PX_Float(pPX, _T("DataQ"), m_fpDataQ,0);
	PX_Short(pPX, _T("ThisPoint"), m_iThisPoint,0);
	PX_Color(pPX, _T("ForeColor"), m_clrForeColor,RGB(0,0,255));

}


/////////////////////////////////////////////////////////////////////////////
// CSmithCtrl::OnResetState - Reset control to default state

void CSmithCtrl::OnResetState()
{ int i;

	COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	// TODO: Reset any other control state here.
	m_bShowAdmittance = FALSE;
	m_iNumPoints = 5;
	m_fpDataM = (float)0.;
	m_fpDataQ = (float)0.;
	m_iThisPoint = 0;
	m_iThisSet=1;
	m_bDataReset = FALSE;
	m_fpVSWRCircle = (float)0.00;
	m_fpQCircle = (float)0.00;
	m_clrForeColor = RGB(0,0,255);
	m_bShowMarker = FALSE;
	m_bShowHistory = TRUE;
	m_bResetAll = FALSE;
	m_fpMarkerM = (float)0.;
	m_fpMarkerQ = (float)0.;
	m_bRedraw = FALSE;
	m_strAbout = "";
	m_bSweep = FALSE;
	m_fpZ0 = (float)50.;
	m_fpCircleM = (float)0.;
	m_fpCircleQ = (float)0.;
	m_fpCircleR = (float)0.;
	m_clrCircleColor = RGB(0,0,0);
	m_iCircleIndex = 0;
	SpotDataM=m_fpDataM;
    SpotDataQ=m_fpDataQ;
	i=m_iNumPoints;
	m_iNumSet = 1;
	if(hGmemM) GlobalFree((GLOBALHANDLE)hGmemM);
    if(hGmemQ) GlobalFree((GLOBALHANDLE)hGmemQ);
	if (i<=0) i = 5;
	    hGmemM=GlobalAlloc(GHND,(i+1)*sizeof(float)*m_iThisSet);    // initial Global Allocation which returns a handle
		hGmemQ=GlobalAlloc(GHND,(i+1)*sizeof(float)*m_iThisSet);
	if (hGmemM==NULL||hGmemQ==NULL){ ThrowError(ERR_MEMALLOC,"Unable to allocate memory.",20001);}
    for ( i = 0 ; i<=CIRCLEMAX; i++)  // iniitalize circle structure
       {
       circlem[i] =0;
       circleq[i] =0;
       circler[i] =0 ;
       circlecolor[i] = 0;
       }



	for ( i = 0 ; i <= DATASETMAX; i++)
	{
		thispointmax[i]=0;  // iniitalize data logs
        forecolor[i] = RGB(0,0,255);
	}


	histindex = 0;
	vswrindex = 0;
	qindex = 0;
	    for ( i = 0 ; i<=HISTORYMAX+1; i++)  // iniitalize historical circle structure
	     {
	       circstor[i].mh =(float)0;
	       circstor[i].qh =(float)0;
	       circstor[i].rh =(float)0 ;
	       circstor[i].cch = 0;

		   vswrstor[i] = (float)0;
		   qstor[i] = (float)0;

		}




}


/////////////////////////////////////////////////////////////////////////////
// CSmithCtrl::AboutBox - Display an "About" box to the user

void CSmithCtrl::AboutBox()
{
	CDialog dlgAbout(IDD_ABOUTBOX_SMITH);
	dlgAbout.DoModal();
}


/////////////////////////////////////////////////////////////////////////////
// CSmithCtrl::PreCreateWindow

BOOL CSmithCtrl::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: Modify windows style flags as necessary
	cs.style &= ~WS_BORDER; // Create a window without border

	return COleControl::PreCreateWindow(cs);
}


/////////////////////////////////////////////////////////////////////////////
// Member functions ported from VBX control

BOOL CSmithCtrl::GetShowAdmittance()
{
	return m_bShowAdmittance;
}

void CSmithCtrl::SetShowAdmittance(BOOL bShowAdmittance)
{ 
	if(m_bShowAdmittance != bShowAdmittance)
	{
		m_bShowAdmittance = bShowAdmittance;
		InvalidateControl();
		// Update any property browser
		BoundPropertyChanged( dispidShowAdmittance );
	}
}




short CSmithCtrl::GetNumPoints()
{
	return m_iNumPoints;
}

void CSmithCtrl::SetNumPoints(short iNumPoints)
{ int i,j,k;

	i=iNumPoints;
	j=GetNumSets();
	if(i < 2 || i >ARRMAX) ThrowError(ERR_OUTRANGE,"NumPoints property is out of range.",20000);
	else if(i*j > 4040 ) ThrowError(ERR_DATAPOINTS,"NumSets x NumPoints property exceeds range.",20008);
	else
	{
	k=GetThisSet();
	j=(SHORT)GlobalSize(hGmemQ);                       //maximum allowable and atleast 2
    j=(SHORT) (j/sizeof(float));                       // get the present size of memory (j)
	  if(i*k  > j)                           // if the amount requested is greater then reallocate movable mem
	    {
	     GlobalReAlloc((GLOBALHANDLE)hGmemM, (i+1)*sizeof(float)*k,GHND);
	     GlobalReAlloc((GLOBALHANDLE)hGmemQ, (i+1)*sizeof(float)*k,GHND);
	     if (hGmemM==NULL||hGmemQ==NULL){ ThrowError(ERR_MEMALLOC,"Unable to allocate memory.",20001); return;}
	    }
		m_iNumPoints = iNumPoints;
		SetModifiedFlag();
		// Update any property browser
	    BoundPropertyChanged( dispidNumPoints );
	}
}

float CSmithCtrl::GetDataM()
{
	return m_fpDataM;
}

void CSmithCtrl::SetDataM(float fpDataM)
{ int j,k;

 j=GetThisPoint();
 if(j >= GetNumPoints()) ThrowError(ERR_THISPOINT,"ThisPoint property exceeds NumPoints.",20003);
 else
 {
  k=GetThisSet();
  m_fpDataM = fpDataM;
  if(!m_bSweep) {SpotDataM = m_fpDataM;	SetModifiedFlag();}  // save spot data
  else
	{
      dataM = (float FAR *)GlobalLock(hGmemM); // get mempointer
      *(dataM + j+((k-1)*thispointmax[k-1])) = m_fpDataM;    // save sweep data
      GlobalUnlock(hGmemM);     // release mempointer
	}
// Update any property browser
	 BoundPropertyChanged( dispidDataM);
 }
}

float CSmithCtrl::GetDataQ()
{
	return m_fpDataQ;
}

void CSmithCtrl::SetDataQ(float fpDataQ)
{ int j,k,i;

j=GetThisPoint();
k=GetThisSet();
i=GetNumPoints();
if(j >= i) ThrowError(ERR_THISPOINT,"ThisPoint property exceeds NumPoints.",20003);
 else
 {
   m_fpDataQ = fpDataQ;
  if(!m_bSweep) SpotDataQ = m_fpDataQ; // save spot data
  else
  {
      dataQ = (float FAR *)GlobalLock(hGmemQ); // get mempointer
      *(dataQ+j +((k-1)*thispointmax[k-1])) = m_fpDataQ;   // save sweep data
      GlobalUnlock(hGmemQ);     // release mempointer
	  ++m_iThisPoint;
	  ++thispointmax[k];
  }

 SetModifiedFlag();
 // Update any property browser
 BoundPropertyChanged( dispidDataQ);
 InvalidateControl();
 }
}

short CSmithCtrl::GetThisPoint()
{
	return m_iThisPoint;
}

void CSmithCtrl::SetThisPoint(short iThisPoint)
{ int i,k;

  i=GetNumPoints();
  if(iThisPoint >= i) ThrowError(ERR_THISPOINT,"ThisPoint property exceeds NumPoints.",20003);
  else
  {
	k=GetThisSet();
    m_iThisPoint = iThisPoint;
   thispointmax[k] = iThisPoint;
    SetModifiedFlag();
	// Update any property browser
   BoundPropertyChanged( dispidThisPoint);
  }
}

short CSmithCtrl::GetThisSet()
{

	return m_iThisSet;
}

void CSmithCtrl::SetThisSet(short nNewValue)
{ int i, k;
  short j;

	i=GetNumPoints();
	k=nNewValue;
		if(k < 1 || k >GetNumSets()) ThrowError(ERR_OUTRANGE1,"ThisSet property is out of range.",20006);
	else
	{
	j=(SHORT)GlobalSize(hGmemQ);                       //maximum allowable and atleast 2
    j=(SHORT) (j/sizeof(float));                       // get the present size of memory (j)
	  if(i*k  > j)                           // if the amount requested is greater then reallocate movable mem
	    {
	     GlobalReAlloc((GLOBALHANDLE)hGmemM, (i+1)*sizeof(float)*k,GHND);
	     GlobalReAlloc((GLOBALHANDLE)hGmemQ, (i+1)*sizeof(float)*k,GHND);
	     if (hGmemM==NULL||hGmemQ==NULL){ ThrowError(ERR_MEMALLOC,"Unable to allocate memory.",20001); return;}
	    }
		m_iThisSet= nNewValue;
		// Update any property browser
      BoundPropertyChanged( dispidThisSet);
	}
}


short CSmithCtrl::GetNumSets()
{

	return m_iNumSet;
}

void CSmithCtrl::SetNumSets(short nNewValue)
{ int k, i;
    k = GetNumPoints();
    i = nNewValue;
	if(i >DATASETMAX || i < 1) ThrowError(ERR_NUMSETS,"NumSets property exceeds range.",20007);
    else if(i*k > 4040 ) ThrowError(ERR_DATAPOINTS,"NumSets x NumPoints property exceeds range.",20008);
    else
	{
	 m_iNumSet=nNewValue;
	 // Update any property browser
     BoundPropertyChanged( dispidNumSets);
	}
}

OLE_COLOR CSmithCtrl::GetForeColor()
{

	return m_clrForeColor;
}

void CSmithCtrl::SetForeColor(OLE_COLOR nNewValue)
{

 k=GetThisSet();
 forecolor[k]=m_clrForeColor=nNewValue;
 SetModifiedFlag();
}


BOOL CSmithCtrl::GetDataReset()
{
	return m_bDataReset;
}

void CSmithCtrl::SetDataReset(BOOL bDataReset)
{int i;

	if(bDataReset)
	{
	m_fpDataM = (float)0.;
	m_fpDataQ = (float)0.;
	SpotDataM=0;
    SpotDataQ=0;
	m_iThisPoint = 0;
	k=m_iThisSet=1;

	m_fpCircleM = (float)0.;
	m_fpCircleQ = (float)0.;
	m_fpCircleR = (float)0.;
	m_clrCircleColor = RGB(0,0,0);
	m_iCircleIndex = 0;

	dataM =(float FAR *)GlobalLock(hGmemM);  // get pointer for data
	dataQ= (float FAR *)GlobalLock(hGmemQ);  // This set is 1
   	for (i=0; i<m_iNumPoints; i++)
     {
       *(dataM + i) = 0;
	   *(dataQ + i) = 0;
     }
     GlobalUnlock((GLOBALHANDLE)hGmemM);  // release pointer
     GlobalUnlock((GLOBALHANDLE)hGmemQ);

   	 for ( i = 0 ; i<=CIRCLEMAX; i++)  // iniitalize circle structure
       {
       circlem[i] =0;
       circleq[i] =0;
       circler[i] =0 ;
       circlecolor[i] = 0;
       }

       // clear historical on reset request
	   histindex = 0;
	   vswrindex = 0;
	   qindex = 0;
	    for ( i = 0 ; i<=HISTORYMAX+1; i++)  // iniitalize historical structures
	       {
	       circstor[i].mh =0;
	       circstor[i].qh =0;
	       circstor[i].rh =0 ;
	       circstor[i].cch = 0;
		   vswrstor[i] = 0;
		   qstor[i] = 0;
       }

	 for ( i = 0 ; i<=DATASETMAX; i++){thispointmax[i]=0;forecolor[i] = 0;}
	 m_iNumSet =1;
	 m_bDataReset =FALSE;
	 InvalidateControl();
// Update any property browser
   BoundPropertyChanged( dispidDataReset);
	}

}

float CSmithCtrl::GetVSWRCircle()
{
	return m_fpVSWRCircle;
}

void CSmithCtrl::SetVSWRCircle(float fpVSWRCircle)
{
	if(m_fpVSWRCircle != fpVSWRCircle && fpVSWRCircle >0)
	{
		m_fpVSWRCircle = fpVSWRCircle;

		InvalidateControl();

       // save for redraw capability
	       vswrindex++;
	       // bounds check roll over to 1 if maxed out
		   // 0 means empty
	       if (vswrindex > HISTORYMAX+1) 
		   { vswrindex = 1; }
           vswrstor[vswrindex] =fpVSWRCircle;

		// Update any property browser
		BoundPropertyChanged( dispidVSWRCircle);
	}
}

float CSmithCtrl::GetQCircle()
{
	return m_fpQCircle;
}

void CSmithCtrl::SetQCircle(float fpQCircle)
{
	if(m_fpQCircle != fpQCircle)
	{
		m_fpQCircle = fpQCircle;
		InvalidateControl();

	       // save for redraw capability
	       qindex++;
	       // bounds check roll over to 1 if maxed out
		   // 0 means empty
	       if (qindex > HISTORYMAX+1) 
		   { qindex = 1; }
            qstor[qindex]=fpQCircle;

		// Update any property browser
        BoundPropertyChanged( dispidQCircle);
	}
}

BOOL CSmithCtrl::GetShowMarker()
{
	return m_bShowMarker;
}

void CSmithCtrl::SetShowMarker(BOOL bShowMarker)
{
	if(m_bShowMarker != bShowMarker)
	{
		m_bShowMarker = bShowMarker;
		InvalidateControl();
		// Update any property browser
		BoundPropertyChanged( dispidShowMarker);
	}
}


BOOL CSmithCtrl::GetShowHistory()
{
	return m_bShowHistory;
}

void CSmithCtrl::SetShowHistory(BOOL bShowHistory)
{// true = show history, false don't show

	    m_bShowHistory = bShowHistory;
		InvalidateControl();
		// Update any property browser
		BoundPropertyChanged( dispidShowHistory);
}

BOOL CSmithCtrl::GetResetAll()
{
	return m_bResetAll;
}

void CSmithCtrl::SetResetAll(BOOL bResetAll)
{
	if(bResetAll)
	{
	 OnResetState();
	 m_bResetAll = FALSE;
	 InvalidateControl();
	 // Update any property browser
	 BoundPropertyChanged( dispidResetAll );
	}

}

float CSmithCtrl::GetMarkerM()
{
	return m_fpMarkerM;
}

void CSmithCtrl::SetMarkerM(float fpMarkerM)
{
	m_fpMarkerM = fpMarkerM;
	// Update any property browser
   BoundPropertyChanged( dispidMarkerM);
}

float CSmithCtrl::GetMarkerQ()
{
	return m_fpMarkerQ;
}

void CSmithCtrl::SetMarkerQ(float fpMarkerQ)
{
	m_fpMarkerQ = fpMarkerQ;
	InvalidateControl();
	// Update any property browser
   BoundPropertyChanged( dispidMarkerQ);
}

BOOL CSmithCtrl::GetRedraw()
{
	return m_bRedraw;
}

void CSmithCtrl::SetRedraw(BOOL bRedraw)
{short i;
   if(bRedraw)
   {
	   InvalidateControl();
	   m_bRedraw = FALSE;

       // clear historical on redraw request
	   histindex = 0;
	   vswrindex = 0;
	   qindex = 0;
	    for ( i = 0 ; i<=HISTORYMAX+1; i++)  // iniitalize historical circle structure
	    {
	       circstor[i].mh =0;
	       circstor[i].qh =0;
	       circstor[i].rh =0 ;
	       circstor[i].cch = 0;

		   vswrstor[i] = 0;
		   qstor[i] = 0;
		}

   // Update any property browser
   BoundPropertyChanged( dispidRedraw);
   }

}


BOOL CSmithCtrl::GetSweep()
{
	return m_bSweep;
}

void CSmithCtrl::SetSweep(BOOL bSweep)
{
	m_bSweep = bSweep;
	// Update any property browser
   BoundPropertyChanged( dispidSweep);
}



float CSmithCtrl::GetZ0()
{
	return m_fpZ0;
}

void CSmithCtrl::SetZ0(float fpZ0)
{

	if(m_fpZ0 != fpZ0 && fpZ0 > 0)
	{
		m_fpZ0 = fpZ0;
		InvalidateControl();
		// Update any property browser
        BoundPropertyChanged( dispidZ0);
	}
}

float CSmithCtrl::GetCircleM()
{
	return m_fpCircleM;
}

void CSmithCtrl::SetCircleM(float fpCircleM)
{short i;
	i=(short)GetCircleIndex();
	circlem[i]=m_fpCircleM = fpCircleM;
	// Update any property browser
    BoundPropertyChanged( dispidCircleM);
}

float CSmithCtrl::GetCircleQ()
{
	return m_fpCircleQ;
}

void CSmithCtrl::SetCircleQ(float fpCircleQ)
{short i;
	i=(short)GetCircleIndex();
	circleq[i]=m_fpCircleQ = fpCircleQ;
	// Update any property browser
    BoundPropertyChanged( dispidCircleQ);
}

float CSmithCtrl::GetCircleR()
{
		return m_fpCircleR;
}

void CSmithCtrl::SetCircleR(float fpCircleR)
{ short i;
	i=(short)GetCircleIndex();
	circler[i]=m_fpCircleR = fpCircleR;
	if(m_fpCircleR>0) 
	{ InvalidateControl();

	       // save for redraw capability
	       histindex++;
	       // bounds check roll over to 1 if maxed out
		   // 0 means empty
	       if (histindex > HISTORYMAX+1) 
		   { histindex = 1; }
           circstor[histindex].mh =circlem[i];
	       circstor[histindex].qh =circleq[i];
	       circstor[histindex].rh =circler[i] ;
	       circstor[histindex].cch = circlecolor[i];
	}

	// Update any property browser
    BoundPropertyChanged( dispidCircleR);
}

OLE_COLOR CSmithCtrl::GetCircleColor()
{
	return m_clrCircleColor;
}

void CSmithCtrl::SetCircleColor(OLE_COLOR clrCircleColor)
{short i;
	i= GetCircleIndex();
	circlecolor[i]=m_clrCircleColor = clrCircleColor;
	// Update any property browser
    BoundPropertyChanged( dispidCircleColor);
}

short CSmithCtrl::GetCircleIndex()
{
	return m_iCircleIndex;
}

void CSmithCtrl::SetCircleIndex(short iCircleIndex)
{ short i;
i=iCircleIndex;
if(i >CIRCLEMAX || i < 0) ThrowError(ERR_CIRCLEINDEX,"CircleIndex property exceeds range.",20002);
else{
	m_fpCircleM =circlem[i];
	m_fpCircleQ =circleq[i];
	m_fpCircleR =circler[i];
	m_clrCircleColor=circlecolor[i];
	m_iCircleIndex = i;
	// Update any property browser
    BoundPropertyChanged( dispidCircleIndex);
	}
}



/////////////////////////////////////////////////////////////////////////////
// CSmithCtrl message handlers

//---------------------------------------------------------------------------
// Return TRUE if the given coordinates are inside of the Smith Chart
//---------------------------------------------------------------------------
BOOL CSmithCtrl::InSmith(float x1, float y1)
{
    double r,x,y;
    x= (double)x1;
    y = (double)y1;
    r = AXIS_RANGE;

    return ((x * x) / (r*r) + (y * y) / (r*r) <= 1);
}



void CSmithCtrl::OnLButtonDblClk(UINT nFlags, CPoint point)
{

	CRect rc;
	float ext;
	float x,y,r1,q;
	int range;
	GetClientRect(rc);
	if (rc.right < rc.bottom)
		{
		ext= (float) rc.right/2;         // MMISOTROPIC
		range =AXIS_RANGE+20;
		}
	else
		{
		ext =(float)rc.bottom/2;
		range =AXIS_RANGE+80;
		}
	x=(float) ((point.x - rc.right/2)*(range/ext));  // get logical coordinates  for x&y
	y =(float) ((point.y- rc.bottom/2)*(range/-ext));
   	r1=(float) sqrt(((x*x)+(y*y))); // convert to polar form
    if(r1==0) q=0 ; else q= (float)(180*asin(y/r1)/Pi);
	if(x < 0 && y >=0) q= 180-q; else if(x<0 && y<0) q= -180-q;  // Adjust for the second and third quadrant
	r1 = r1/AXIS_RANGE;
	if(InSmith(x,y)) FireClickIn(&r1,&q);          // fire a windows event
	else FireClickOut();

    COleControl::OnLButtonDblClk(nFlags, point);
}

void CSmithCtrl::OnMouseMove(UINT nFlags, CPoint point)
{
    CRect rc;
	float ext;
	float x,y,r1,q;
	int range;
	GetClientRect(rc);
	if (rc.right < rc.bottom)
		{
		ext= (float) rc.right/2;         // MMISOTROPIC
		range =AXIS_RANGE+20;
		}
	else
		{
		ext =(float)rc.bottom/2;
		range =AXIS_RANGE+80;
		}
	x=(float) ((point.x - rc.right/2)*(range/ext));  // get logical coordinates  for x&y
	y =(float) ((point.y- rc.bottom/2)*(range/-ext));
   	r1=(float) sqrt(((x*x)+(y*y))); // convert to polar form
    if(r1==0) q=0 ; else q= (float)(180*asin(y/r1)/Pi);
	if(x < 0 && y >=0) q= 180-q; else if(x<0 && y<0) q= -180-q;  // Adjust for the second and third quadrant
	r1 = r1/AXIS_RANGE;

    if(InSmith(x,y))FireMouseMove(&r1,&q);          // fire a windows event
    else
    {
	  r1 = 2; q=0;
	  FireMouseMove(&r1,&q); //send a sign if the mousemove was detected outside og the smithchart
	}
	COleControl::OnMouseMove(nFlags, point);
}

