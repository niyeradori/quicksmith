// smithppg.h : Declaration of the CSmithPropPage property page class.

////////////////////////////////////////////////////////////////////////////
// CSmithPropPage : See smithppg.cpp for implementation.

class CSmithPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CSmithPropPage)
	DECLARE_OLECREATE_EX(CSmithPropPage)

// Constructor
public:
	CSmithPropPage();

// Dialog Data
	//{{AFX_DATA(CSmithPropPage)
	enum { IDD = IDD_PROPPAGE_SMITH };
	BOOL	m_showAdmittance;
	float	m_vswrcircle;
	float	m_qcircle;
	float	m_markerm;
	float	m_markerq;
	float	m_zo;
	BOOL	m_showmarker;
	BOOL	m_resetall;
	BOOL	m_redraw;
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	//{{AFX_MSG(CSmithPropPage)
	afx_msg void OnResetall();
	afx_msg void OnRedraw();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

};
