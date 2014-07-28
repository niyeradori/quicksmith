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
// smithctl.h : Declaration of the CSmithCtrl OLE control class.


#define Pi 3.141592654
#define BWIDTH 11     //Bitmap width and height for the marker sprite
#define BHEIGTH	11    // use odd numbers to center remember it starts at zero
#define AXIS_RANGE 1000
#define ARRMAX 1010
#define CIRCLEMAX 10
#define HISTORYMAX 30  //save the last 30 circles on looping basis
#define DATASETMAX 10
#define Circle(X,Y,R) Ellipse(X-R,Y+R,X+R,Y-R)
#define MyArc(X,Y,R,A1,A2) Arc(X-R,Y+R,X+R,Y-R,X+(short)(R*cos(Pi*A1/180)),Y+(short)(R*sin(Pi*A1/180)),X+(short)(R*cos(Pi*A2/180)),Y+(short)(R*sin(Pi*A2/180)))
#define MyQArc1(X,Y,R) Arc(X-R,Y+R,X+R,Y-R,-AXIS_RANGE,0,AXIS_RANGE,0)// +Yaxis arc starting from -r,0 to r,0
#define MyQArc2(X,Y,R) Arc(X-R,Y+R,X+R,Y-R,AXIS_RANGE,0,-AXIS_RANGE,0)// -Yaxis arc starting from -r,0 to r,0
#define MyRectangle(X,Y,R) Rectangle(X-R,Y+R,X+R,Y-R)
#define ERR_OUTRANGE CUSTOM_CTL_SCODE(20000)
#define ERR_MEMALLOC CUSTOM_CTL_SCODE(20001)
#define ERR_CIRCLEINDEX CUSTOM_CTL_SCODE(20002)
#define ERR_THISPOINT CUSTOM_CTL_SCODE(20003)
#define ERR_BADINDX CUSTOM_CTL_SCODE(20004)
#define ERR_GAMMAOUTRANGE CUSTOM_CTL_SCODE(20005)
#define ERR_OUTRANGE1 CUSTOM_CTL_SCODE(20006)
#define ERR_NUMSETS CUSTOM_CTL_SCODE(20007)
#define ERR_DATAPOINTS CUSTOM_CTL_SCODE(20007)
/////////////////////////////////////////////////////////////////////////////
// CSmithCtrl : See smithctl.cpp for implementation.

class CSmithCtrl : public COleControl
{
	DECLARE_DYNCREATE(CSmithCtrl)

// Constructor
public:
	CSmithCtrl();

// Overrides

	// Drawing function
	virtual void OnDrawMetafile(CDC* pdc, const CRect& rcBounds);

	virtual void OnDraw(
				CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);

	// Persistence
	virtual void DoPropExchange(CPropExchange* pPX);

	// Reset control state
	virtual void OnResetState();

// Implementation
protected:
	~CSmithCtrl();

typedef struct CHistory
   { float mh;
		 float qh;
		 float rh;
		 OLE_COLOR cch;
	}CHistory;

    CBitmap* mHBMImage;
	CBitmap* mHBMMask;
	BOOL first_draw;
	int i,j,k;
	GLOBALHANDLE hGmemM;
	GLOBALHANDLE hGmemQ;
	float FAR *dataM;
	float FAR *dataQ;  // sweep data pointers
	float SpotDataM;
	float SpotDataQ;

	float circlem [CIRCLEMAX+1]; // constcircle parameters
  float circleq [CIRCLEMAX+1];
  float circler [CIRCLEMAX+1];
  OLE_COLOR circlecolor[CIRCLEMAX+1];

  CHistory circstor[HISTORYMAX+2];  // added to store historical circles for display
  short histindex;
  float qstor[HISTORYMAX+2];        // added to store historical circles for display
  short qindex;
  float vswrstor[HISTORYMAX+2];     // added to store historical circles for display
  short vswrindex;

	int thispointmax[DATASETMAX];
	void MapSmith(CDC* pdc, CRect rc);  //  Sets client coordinates
	void DrawSmith(CDC* pdc,CRect rc );
	BOOL CSmithCtrl::InSmith(float x1,float y1);
	OLE_COLOR forecolor[DATASETMAX+1];

    BEGIN_OLEFACTORY(CSmithCtrl)        // Class factory and guid
		virtual BOOL VerifyUserLicense();
		virtual BOOL GetLicenseKey(DWORD, BSTR FAR*);
	END_OLEFACTORY(CSmithCtrl)

	//DECLARE_OLECREATE_EX(CSmithCtrl)    // Class factory and guid
	DECLARE_OLETYPELIB(CSmithCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CSmithCtrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CSmithCtrl)		// Type name and misc status

	// VBX port support
	BOOL PreCreateWindow(CREATESTRUCT& cs);

	// Storage for Get/Set properties
	BOOL m_bShowAdmittance;
	short m_iNumPoints;
	float m_fpDataM;
	float m_fpDataQ;
	short m_iThisPoint;
	BOOL m_bDataReset;
	float m_fpVSWRCircle;
	float m_fpQCircle;
	BOOL m_bShowMarker;
	BOOL m_bShowHistory;
	BOOL m_bResetAll;
	float m_fpMarkerM;
	float m_fpMarkerQ;
	BOOL m_bRedraw;
	CString m_strAbout;
	BOOL m_bSweep;
	float m_fpZ0;
	float m_fpCircleM;
	float m_fpCircleQ;
	float m_fpCircleR;
	OLE_COLOR m_clrCircleColor;
	short m_iCircleIndex;
	short m_iThisSet;
	OLE_COLOR m_clrForeColor;
	short m_iNumSet;

// Message maps
	//{{AFX_MSG(CSmithCtrl)
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

// Dispatch maps
	//{{AFX_DISPATCH(CSmithCtrl)
	BOOL m_showhistory;
	afx_msg BOOL GetShowAdmittance();
	afx_msg void SetShowAdmittance(BOOL bNewValue);
	afx_msg short GetNumPoints();
	afx_msg void SetNumPoints(short nNewValue);
	afx_msg float GetDataM();
	afx_msg void SetDataM(float newValue);
	afx_msg float GetDataQ();
	afx_msg void SetDataQ(float newValue);
	afx_msg short GetThisPoint();
	afx_msg void SetThisPoint(short nNewValue);
	afx_msg BOOL GetDataReset();
	afx_msg void SetDataReset(BOOL bNewValue);
	afx_msg float GetVSWRCircle();
	afx_msg void SetVSWRCircle(float newValue);
	afx_msg float GetQCircle();
	afx_msg void SetQCircle(float newValue);
	afx_msg BOOL GetShowMarker();
	afx_msg BOOL GetShowHistory();
	afx_msg void SetShowMarker(BOOL bNewValue);
    afx_msg void SetShowHistory(BOOL bShowHistory);
	afx_msg BOOL GetResetAll();
	afx_msg void SetResetAll(BOOL bNewValue);
	afx_msg float GetMarkerM();
	afx_msg void SetMarkerM(float newValue);
	afx_msg float GetMarkerQ();
	afx_msg void SetMarkerQ(float newValue);
	afx_msg BOOL GetRedraw();
	afx_msg void SetRedraw(BOOL bNewValue);
	afx_msg BOOL GetSweep();
	afx_msg void SetSweep(BOOL bNewValue);
	afx_msg float GetZ0();
	afx_msg void SetZ0(float newValue);
	afx_msg float GetCircleM();
	afx_msg void SetCircleM(float newValue);
	afx_msg float GetCircleQ();
	afx_msg void SetCircleQ(float newValue);
	afx_msg float GetCircleR();
	afx_msg void SetCircleR(float newValue);
	afx_msg OLE_COLOR GetCircleColor();
	afx_msg void SetCircleColor(OLE_COLOR nNewValue);
	afx_msg short GetCircleIndex();
	afx_msg void SetCircleIndex(short nNewValue);
	afx_msg short GetThisSet();
	afx_msg void SetThisSet(short nNewValue);
	afx_msg short GetNumSets();
	afx_msg void SetNumSets(short nNewValue);
	afx_msg OLE_COLOR GetForeColor();
	afx_msg void SetForeColor(OLE_COLOR nNewValue);
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()

	afx_msg void AboutBox();

// Event maps
	//{{AFX_EVENT(CSmithCtrl)
	void FireClickIn(float FAR* M, float FAR* Q)
		{FireEvent(eventidClickIn,EVENT_PARAM(VTS_PR4  VTS_PR4), M, Q);}
	void FireClickOut()
		{FireEvent(eventidClickOut,EVENT_PARAM(VTS_NONE));}
	void FireMouseMove(float FAR* M, float FAR* Q)
		{FireEvent(eventidMouseMove,EVENT_PARAM(VTS_PR4  VTS_PR4), M, Q);}
	//}}AFX_EVENT
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	enum {
	//{{AFX_DISP_ID(CSmithCtrl)
	dispidShowAdmittance = 2L,
	dispidNumPoints = 3L,
	dispidDataM = 4L,
	dispidDataQ = 5L,
	dispidThisPoint = 6L,
	dispidDataReset = 7L,
	dispidVSWRCircle = 8L,
	dispidQCircle = 9L,
	dispidShowMarker = 10L,
	dispidResetAll = 11L,
	dispidMarkerM = 12L,
	dispidMarkerQ = 13L,
	dispidRedraw = 14L,
	dispidSweep = 15L,
	dispidZ0 = 16L,
	dispidCircleM = 17L,
	dispidCircleQ = 18L,
	dispidCircleR = 19L,
	dispidCircleColor = 20L,
	dispidCircleIndex = 21L,
	dispidThisSet = 22L,
	dispidNumSets = 23L,
	dispidShowHistory = 1L,
	eventidClickIn = 1L,
	eventidClickOut = 2L,
	eventidMouseMove = 3L,
	//}}AFX_DISP_ID
	};
};
