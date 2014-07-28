// Datappg.cpp : implementation file
//

#include "stdafx.h"
#include "smith.h"
#include "Datappg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CDataPropPage dialog

IMPLEMENT_DYNCREATE(CDataPropPage, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CDataPropPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CDataPropPage)

	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

// {75E627A4-3911-11D1-9376-C8DD03C10000}
IMPLEMENT_OLECREATE_EX(CDataPropPage, "Smith.CDataPropPage",
	0x75e627a4, 0x3911, 0x11d1, 0x93, 0x76, 0xc8, 0xdd, 0x3, 0xc1, 0x0, 0x0)


/////////////////////////////////////////////////////////////////////////////
// CDataPropPage::CDataPropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CDataPropPage

BOOL CDataPropPage::CDataPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Define string resource for page type; replace '0' below with ID.

	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_DATA_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CDataPropPage::CDataPropPage - Constructor

// TODO: Define string resource for page caption; replace '0' below with ID.

CDataPropPage::CDataPropPage() :
	COlePropertyPage(IDD, IDS_DATA_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CDataPropPage)
	m_circleindex = 0.0f;
	m_circlem = 0.0f;
	m_circleq = 0.0f;
	m_circler = 0.0f;
	m_datam = 0.0f;
	m_dataq = 0.0f;
	m_numpoints = 0.0f;
	m_numsets = 0.0f;
	m_thsipoint = 0.0f;
	m_thisset = 0.0f;
	m_sweep = FALSE;
	m_datareset = FALSE;
	//}}AFX_DATA_INIT
}


/////////////////////////////////////////////////////////////////////////////
// CDataPropPage::DoDataExchange - Moves data between page and properties

void CDataPropPage::DoDataExchange(CDataExchange* pDX)
{
	// NOTE: ClassWizard will add DDP, DDX, and DDV calls here
	//    DO NOT EDIT what you see in these blocks of generated code !
	//{{AFX_DATA_MAP(CDataPropPage)
	DDP_Text(pDX, IDC_CIRCLEINDEX, m_circleindex, _T("CircleIndex") );
	DDX_Text(pDX, IDC_CIRCLEINDEX, m_circleindex);
	DDP_Text(pDX, IDC_CIRCLEM, m_circlem, _T("CircleM") );
	DDX_Text(pDX, IDC_CIRCLEM, m_circlem);
	DDP_Text(pDX, IDC_CIRCLEQ, m_circleq, _T("CircleQ") );
	DDX_Text(pDX, IDC_CIRCLEQ, m_circleq);
	DDP_Text(pDX, IDC_CIRCLER, m_circler, _T("CircleR") );
	DDX_Text(pDX, IDC_CIRCLER, m_circler);
	DDP_Text(pDX, IDC_DATAM, m_datam, _T("DataM") );
	DDX_Text(pDX, IDC_DATAM, m_datam);
	DDP_Text(pDX, IDC_DATAQ, m_dataq, _T("DataQ") );
	DDX_Text(pDX, IDC_DATAQ, m_dataq);
	DDP_Text(pDX, IDC_NUMPOINTS, m_numpoints, _T("NumPoints") );
	DDX_Text(pDX, IDC_NUMPOINTS, m_numpoints);
	DDP_Text(pDX, IDC_NUMSETS, m_numsets, _T("NumSets") );
	DDX_Text(pDX, IDC_NUMSETS, m_numsets);
	DDP_Text(pDX, IDC_THISPOINT, m_thsipoint, _T("ThisPoint") );
	DDX_Text(pDX, IDC_THISPOINT, m_thsipoint);
	DDP_Text(pDX, IDC_THISSET, m_thisset, _T("ThisSet") );
	DDX_Text(pDX, IDC_THISSET, m_thisset);
	DDP_Check(pDX, IDC_SWEEP, m_sweep, _T("Sweep") );
	DDX_Check(pDX, IDC_SWEEP, m_sweep);
	DDP_Check(pDX, IDC_DATARESET, m_datareset, _T("DataReset") );
	DDX_Check(pDX, IDC_DATARESET, m_datareset);
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);
}

