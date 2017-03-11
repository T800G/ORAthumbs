#ifndef _THUMBNAILHANDLER_H__5970152C_B024_4304_8C88_346E588DAD65__INCLUDED_
#define _THUMBNAILHANDLER_H__5970152C_B024_4304_8C88_346E588DAD65__INCLUDED_
#pragma once

#ifndef STRICT
#define STRICT
#endif

#include <shlObj.h>
#pragma comment(lib,"shlwapi.lib")

#include "orathumbnail.h"
#include "resource.h"


/////////////////////////////////////////////////////////////////////////////
// CORAThumbnail

class ATL_NO_VTABLE CThumbnailHandler : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CThumbnailHandler, &CLSID_ORAThumbnail>,
	public IPersistFile,
	public IExtractImage2
{
public:
	CThumbnailHandler()
	{
		m_thumbSize.cx=0;
		m_thumbSize.cy=0;
	}

	BEGIN_COM_MAP(CThumbnailHandler)
		COM_INTERFACE_ENTRY(IPersistFile)
		COM_INTERFACE_ENTRY(IExtractImage)
		COM_INTERFACE_ENTRY(IExtractImage2)
	END_COM_MAP()

	DECLARE_NOT_AGGREGATABLE(CThumbnailHandler) 
	DECLARE_REGISTRY_RESOURCEID(IDR_ORAThumbnail)
	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct(void)
	{
		ATLTRACE("CThumbnailHandler::FinalConstruct\n");
	return S_OK;
	}
	void FinalRelease(void)
	{
		ATLTRACE("CThumbnailHandler::FinalRelease\n");
	}

	// IPersistFile
	STDMETHOD(GetClassID)(LPCLSID clsid){return E_NOTIMPL;}
	STDMETHOD(IsDirty)(VOID){return E_NOTIMPL;}
	STDMETHOD(Save)(LPCOLESTR, BOOL){return E_NOTIMPL;}
	STDMETHOD(SaveCompleted)(LPCOLESTR){return E_NOTIMPL;}
	STDMETHOD(GetCurFile)(LPOLESTR FAR*){return E_NOTIMPL;}
	STDMETHOD(Load)(LPCOLESTR wszFile, DWORD dwMode)
	{
		ATLTRACE("IPersistFile::Load\n");
	return m_thumb.Initialize(wszFile);
	}

	// IExtractImage
	STDMETHOD(GetLocation)(LPWSTR pszPathBuffer, DWORD cchMax, DWORD *pdwPriority,
							const SIZE *prgSize, DWORD dwRecClrDepth, DWORD *pdwFlags)
	{
		ATLTRACE("IExtractImage2::GetLocation\n");
		HRESULT hr=S_OK;
		if (pszPathBuffer) hr=m_thumb.GetLocation(pszPathBuffer, cchMax);
		HR_TRACE(hr);

		m_thumbSize.cx=prgSize->cx;
		m_thumbSize.cy=prgSize->cy;
		*pdwFlags |= (IEIFLAG_CACHE | IEIFLAG_REFRESH);//cache thumbnails, enable refresh option (XP)
#ifdef _DEBUG
		if (*pdwFlags & IEIFLAG_ASPECT) ATLTRACE("IExtractImage::GetLocation : IEIFLAG_ASPECT flag set\n");
		if (*pdwFlags & IEIFLAG_ASYNC) ATLTRACE("IExtractImage::GetLocation : IEIFLAG_ASYNC flag set\n");
#endif
	return S_OK;
	}

	STDMETHOD(Extract)(HBITMAP* phBmpThumbnail)
	{
		ATLTRACE("IExtractImage::Extract\n");
	return m_thumb.GetThumbnail(phBmpThumbnail, &m_thumbSize);
	}

	// IExtractImage2
	STDMETHOD(GetDateStamp)(FILETIME *pDateStamp)
	{
		ATLTRACE("IExtractImage2::GetDateStamp\n");
	return m_thumb.GetDateStamp(pDateStamp);
	}

private:
	CORAThumbnail m_thumb;
	SIZE m_thumbSize;
};

#endif//_ORATHUMBNAIL_H__5970152C_B024_4304_8C88_346E588DAD65__INCLUDED_
