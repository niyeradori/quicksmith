// smithppg.cpp : Implementation of the CSmithPropPage property page class.

#include "stdafx.h"
#include "smith.h"
#include "smithppg.h"

#ifdef _DEBUG
#undef THIS_FILE
static char BASED_CODE THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CSmithPropPage, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CSmithPropPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CSmithPropPage)

	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CSmithPropPage, "SMITH.SmithPropPage.1",
	0x5ae352a4, 0xe9a4, 0x11d0, 0x93, 0x6d, 0x44, 0x45, 0x53, 0x54, 0x0, 0x0)


/////////////////////////////////////////////////////////////////////////////
// CSmithPropPage::CSmithPropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CSmithPropPage

BOOL CSmithPropPage::CSmithPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_SMITH_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CSmithPropPage::CSmithPropPage - Constructor

CSmithPropPage::CSmithPropPage() :
	COlePropertyPage(IDD, IDS_SMITH_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CSmithPropPage)
	m_showAdmittance = FALSE;
	m_vswrcircle = 0.0f;
	m_qcircle = 0.0f;
	m_markerm = 0.0f;
	m_markerq = 0.0f;
	m_zo = 0.0f;
	m_showmarker = FALSE;
	m_resetall = FALSE;
	m_redraw = FALSE;
	//}}AFX_DATA_INIT
}


/////////////////////////////////////////////////////////////////////////////
// CSmithPropPage::DoDataExchange - Moves data between page and properties

void CSmithPropPage::DoDataExchange(CDataExchange* pDX)
{
	//{{AFX_DATA_MAP(CSmithPropPage)
	DDP_Check(pDX, IDC_SHOWADMITTANCE, m_showAdmittance, _T("ShowAdmittance") );
	DDX_Check(pDX, IDC_SHOWADMITTANCE, m_showAdmittance);
	DDP_Text(pDX, IDC_VSWRCIRCLE, m_vswrcircle, _T("VSWRCircle") );
	DDX_Text(pDX, IDC_VSWRCIRCLE, m_vswrcircle);
	DDP_Text(pDX, IDC_QCIRCLE, m_qcircle, _T("QCircle") );
	DDX_Text(pDX, IDC_QCIRCLE, m_qcircle);
	DDP_Text(pDX, IDC_MARKERM, m_markerm, _T("MarkerM") );
	DDX_Text(pDX, IDC_MARKERM, m_markerm);
	DDP_Text(pDX, IDC_MARKERQ, m_markerq, _T("MarkerQ") );
	DDX_Text(pDX, IDC_MARKERQ, m_markerq);
	DDP_Text(pDX, IDC_Z0, m_zo, _T("Z0") );
	DDX_Text(pDX, IDC_Z0, m_zo);
	DDP_Check(pDX, IDC_SHOWMARKER, m_showmarker, _T("ShowMarker") );
	DDX_Check(pDX, IDC_SHOWMARKER, m_showmarker);
	DDP_Check(pDX, IDCRESETALL, m_resetall, _T("ResetAll") );
	DDX_Check(pDX, IDCRESETALL, m_resetall);
	DDP_Check(pDX, IDCREDRAW, m_redraw, _T("Redraw") );
	DDX_Check(pDX, IDCREDRAW, m_redraw);
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);
}



