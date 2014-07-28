#if !defined(AFX_DATAPPG_H__75E627A5_3911_11D1_9376_C8DD03C10000__INCLUDED_)
#define AFX_DATAPPG_H__75E627A5_3911_11D1_9376_C8DD03C10000__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000
// Datappg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CDataPropPage : Property page dialog

class CDataPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CDataPropPage)
	DECLARE_OLECREATE_EX(CDataPropPage)

// Constructors
public:
	CDataPropPage();

// Dialog Data
	//{{AFX_DATA(CDataPropPage)
	enum { IDD = IDD_PROPPAGE_DATA };
	float	m_circleindex;
	float	m_circlem;
	float	m_circleq;
	float	m_circler;
	float	m_datam;
	float	m_dataq;
	float	m_numpoints;
	float	m_numsets;
	float	m_thsipoint;
	float	m_thisset;
	BOOL	m_sweep;
	BOOL	m_datareset;
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);        // DDX/DDV support

// Message maps
protected:
	//{{AFX_MSG(CDataPropPage)
	afx_msg void OnDatareset();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DATAPPG_H__75E627A5_3911_11D1_9376_C8DD03C10000__INCLUDED_)
