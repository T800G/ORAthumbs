#ifndef _ORATHUMBNAIL_F029ABF6_5C6A_4664_ACF3_C95FEB754C03_
#define _ORATHUMBNAIL_F029ABF6_5C6A_4664_ACF3_C95FEB754C03_

//#include "gdiplus.h" // uncomment if needed
//#pragma comment(lib,"gdiplus.lib")

#include <Shellapi.h>
#pragma comment(lib,"shell32.lib")
#include <Shlobj.h>

//ATL headers
//#include <atlstr.h> // uncomment if needed
#include <atlimage.h> //must include before strsafe.h
#ifndef _NTFS_MAX_PATH
#define _NTFS_MAX_PATH   32768
#endif
#define _MEM_MAXBUFFER_SIZE 33554432 //32mb

#include <strsafe.h>

#include <comdef.h>
//double expand trick
#define S(x) #x
#define S_(x) S(x)
#define S__LINE__ S_(__LINE__)
#define S__FILE__ S_(__FILE__)
//use S__LINE__ instead of __LINE__
#ifdef _DEBUG
#define HR_TRACE(x) _HR_TRACE(x,S__FILE__,__LINE__)
static void _HR_TRACE(HRESULT hr, LPCSTR szfile, int i=0)
{
	if SUCCEEDED(hr) return;
	USES_CONVERSION;
	_com_error err(hr);
	static TCHAR pbuf[1024];
	StringCchPrintf(pbuf, 1024, L"%ws::%d\nHRESULT 0x%08X\n%ws\n", A2W(szfile), i, hr, err.ErrorMessage());
	//MessageBox(g_hDlg, pbuf, L"HR_TRACE",MB_OK | MB_ICONWARNING);
	ATLTRACE(pbuf);
}
#else
#define HR_TRACE(x)
#endif


typedef struct tagCoTaskMemAutofree
{
	tagCoTaskMemAutofree(void** pp): m_pp(pp) {}
	~tagCoTaskMemAutofree()
	{
		::CoTaskMemFree(*m_pp);
		*m_pp=NULL;//while() should get empty ptr
		//ATLTRACE("CCoTaskMemAutofree::~CCoTaskMemAutofree\n");
	}
private:
	void** m_pp;
} CCoTaskMemAutofree;


//?system thumbextractor
inline HRESULT ThumbnailFromIStream(IStream* pIs, HBITMAP* phBmpThumbnail, const LPSIZE pThumbSize)
{
	*phBmpThumbnail=NULL;
	ATLASSERT(pIs);
	ATL::CImage ci;//uses GDI+ internally
	HRESULT hr=ci.Load(pIs);
	HR_TRACE(hr);
	if (S_OK!=hr) return hr;

	//check size
	int tw=ci.GetWidth();
	int th=ci.GetHeight();
	float rx=(float)pThumbSize->cx/(float)tw;
	float ry=(float)pThumbSize->cy/(float)th;

	//if bigger size
	if ((rx<1) || (ry<1))
	{
		HDC hdcNew=::CreateCompatibleDC(NULL);
		if (NULL==hdcNew) return HRESULT_FROM_WIN32(::GetLastError());

		::SetStretchBltMode(hdcNew, HALFTONE);
		::SetBrushOrgEx(hdcNew, 0,0, NULL);

		//variables retain values until assignment
		tw=(int)(min(rx,ry)*tw);//Warning C424 workaround
		th=(int)(min(rx,ry)*th);

		HBITMAP hbmpNew=::CreateCompatibleBitmap(ci.GetDC(), tw,th);
		ci.ReleaseDC();//don't forget!
		if (hbmpNew)
		{
			HBITMAP hbmpOld=(HBITMAP)::SelectObject(hdcNew, hbmpNew);
			RECT rc={0,0,tw,th};
			::FillRect(hdcNew, &rc, (HBRUSH)::GetStockObject(WHITE_BRUSH));//try HOLLOW_BRUSH ?alpha
			hr=ci.Draw(hdcNew, 0,0, tw,th, 0,0, ci.GetWidth(),ci.GetHeight()) ? S_OK : S_FALSE;
			hbmpNew=(HBITMAP)::SelectObject(hdcNew, hbmpOld);
		}
		else hr=HRESULT_FROM_WIN32(::GetLastError());

		::DeleteDC(hdcNew);
		if (S_OK==hr) *phBmpThumbnail=hbmpNew;
#ifdef _DEBUG
		else ATLASSERT(NULL==hbmpNew);
#endif
	return hr;
	}

	*phBmpThumbnail=ci.Detach();
return hr;
}


class CORAThumbnail
{
public:
	CORAThumbnail(): m_pStg(NULL) {}
	virtual ~CORAThumbnail()
	{
		Release();
	}

	//HRESULT Initialize(IDataObject*)
	HRESULT Initialize(LPCWSTR szFile)
	{
		//Release();
		ATLASSERT(NULL==m_pStg);
		HRESULT hr=CreateInstance();
		HR_TRACE(hr);
		if (S_OK!=hr) return hr;

		LPITEMIDLIST pidl;
		hr=SHParseDisplayName(szFile, NULL, &pidl, 0, NULL);
		HR_TRACE(hr);
		if (S_OK==hr)
		{
			IPersistFolder* pPersistFolder;
			hr=m_pStg->QueryInterface(IID_IPersistFolder, (void**)&pPersistFolder);
			HR_TRACE(hr);
			if(S_OK==hr)
			{
				hr=pPersistFolder->Initialize(pidl);
				pPersistFolder->Release();
			}
			::CoTaskMemFree(pidl);
		}
		HR_TRACE(hr);
	return hr;
	}

	ULONG Release()
	{
		ULONG l=0;
		if (m_pStg) l=m_pStg->Release();
		ATLASSERT(0==l);
		m_pStg=NULL;
		return l;
	}

	HRESULT GetThumbnail(HBITMAP* phBmpThumbnail, const LPSIZE pThumbSize)
	{
		ATLASSERT(m_pStg);
		return ThumbnailFromIStorage(m_pStg, phBmpThumbnail, pThumbSize);
	}

	HRESULT GetLocation(LPWSTR pszPathBuffer, DWORD cchMax)
	{
		ATLASSERT(m_pStg);
		ATLASSERT(pszPathBuffer && cchMax);
		if (NULL==m_pStg) return E_UNEXPECTED;
		if ((NULL==pszPathBuffer) || (0==cchMax)) return E_INVALIDARG;
		STATSTG stat;
		HRESULT hr=m_pStg->Stat(&stat, STATFLAG_DEFAULT);
		HR_TRACE(hr);
		if (S_OK==hr)
		{
			hr=StringCchCopy(pszPathBuffer, cchMax, stat.pwcsName);
			::CoTaskMemFree(stat.pwcsName);
		}
		HR_TRACE(hr);
	return hr;
	}

	HRESULT GetDateStamp(FILETIME *pDateStamp)
	{
		ATLASSERT(m_pStg);
		ATLASSERT(pDateStamp);
		if (NULL==m_pStg) return E_UNEXPECTED;
		if (NULL==pDateStamp) return E_INVALIDARG;
		STATSTG stat;
		HRESULT hr=m_pStg->Stat(&stat, STATFLAG_NONAME);
		HR_TRACE(hr);
		if (S_OK==hr) *pDateStamp=stat.mtime;
	return hr;
	}

private:
	IStorage* m_pStg;

	HRESULT CreateInstance()
	{
		ATLASSERT(NULL==m_pStg);
		const CLSID CLSID_ZipStorageHandler={0xe88dcce0, 0xb7b3, 0x11d1, {0xa9, 0xf0, 0x00, 0xaa, 0x00, 0x60, 0xfa, 0x31}};
		return ::CoCreateInstance(CLSID_ZipStorageHandler, NULL, CLSCTX_INPROC_SERVER, IID_IStorage, (void**)&m_pStg);
	}

	HRESULT ThumbnailFromIStorage(IStorage* pStg, HBITMAP* phBmpThumbnail, const LPSIZE pThumbSize)
	{
		*phBmpThumbnail=NULL;
		CComPtr<IEnumSTATSTG> pEnumSTATSTG;
		HRESULT hr=pStg->EnumElements(NULL, NULL, NULL, &pEnumSTATSTG);
		HR_TRACE(hr);
		if (S_OK==hr)
		{
			STATSTG statstg;
			while (S_OK==(hr=pEnumSTATSTG->Next(1, &statstg, NULL)))
			{
				CCoTaskMemAutofree ctm((void**)&statstg.pwcsName);
				//ATLTRACE(L"statstg.pwcsName=%ws\n", statstg.pwcsName);
				if ((STGTY_STORAGE==statstg.type) && (StrCmpI(statstg.pwcsName, _T("Thumbnails"))==0))//Thumbnails folder
				{
					CComPtr<IStorage> pStgNew;
					hr=pStg->OpenStorage(statstg.pwcsName, NULL, STGM_READ|STGM_SHARE_DENY_WRITE, 0, 0, &pStgNew);
					HR_TRACE(hr);
					if (S_OK==hr)
						hr=ThumbnailFromIStorage(pStgNew, phBmpThumbnail, pThumbSize);
					if (S_FALSE==hr) break;
				}
				else
				if ((STGTY_STREAM==statstg.type) && (StrCmpI(statstg.pwcsName,_T("thumbnail.png"))==0))//thumbnail file
				{
					if (statstg.cbSize.QuadPart>_MEM_MAXBUFFER_SIZE) return E_OUTOFMEMORY;//size check!
					CComPtr<IStream> pStream;
					hr=pStg->OpenStream(statstg.pwcsName, NULL, STGM_READ|STGM_SHARE_DENY_WRITE, NULL, &pStream);
					HR_TRACE(hr);
					if (S_OK==hr)
					{
						//returned pStream is actually ISequentialStream but it works anyway
						hr=ThumbnailFromIStream(pStream, phBmpThumbnail, pThumbSize);
						HR_TRACE(hr);
						hr=(SUCCEEDED(hr) && *phBmpThumbnail) ? S_FALSE : E_FAIL;
						break;
					}
				}
			}//while()
			HR_TRACE(hr);
		}
	return hr;
	}

};


#endif//_ORATHUMBNAIL_F029ABF6_5C6A_4664_ACF3_C95FEB754C03_
