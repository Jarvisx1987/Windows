#include "StdAfx.h"
#include "WindowsThumbnail.h"

// Common
#include "StringUtil.h"

#define THUMB_CX 64

typedef HRESULT (*PFN_SHCreateItemFromParsingName)(__in PCWSTR pszPath, __in_opt IBindCtx *pbc, __in REFIID riid, __deref_out void **ppv);
PFN_SHCreateItemFromParsingName g_pfnSHCreateItemFromParsingName = NULL;

CWindowsThumbnail::CWindowsThumbnail(void) :
	m_pShellItem(NULL),
	m_pThumbProvider(NULL),
	m_pSharedBitmap(NULL),
	m_pThumbnailCache(NULL),
	m_pShellItemImageFactory(NULL),
	m_hBmp(NULL)
{
	HMODULE hDll = GetModuleHandleW(L"shell32.dll");
	if (hDll != NULL)
	{
		g_pfnSHCreateItemFromParsingName = (PFN_SHCreateItemFromParsingName)GetProcAddress(hDll, "SHCreateItemFromParsingName");
	}
}

CWindowsThumbnail::~CWindowsThumbnail(void)
{
	Release();
}

HRESULT CWindowsThumbnail::GetShellThumbnailImageVistaDown(LPCWSTR pwszFilePath, HBITMAP* phBitmap)
{
	START("CWindowsThumbnail::GetShellThumbnailImageVistaDown");
	HRESULT hr = E_FAIL;

	LPITEMIDLIST pidlItems = NULL, pidlURL = NULL, pidlWorkDir = NULL;
	WCHAR awszFilePath[MAX_PATH], awszFileName[MAX_PATH];

	do
	{
		if (pwszFilePath == NULL)
		{
			break;
		}

		if (phBitmap == NULL)
		{
			break;
		}

		wstring wsFilePath = CStringUtil::GetPath(pwszFilePath);
		wcscpy(awszFilePath, wsFilePath.c_str());

		wstring wsFileName = CStringUtil::GetFileName(pwszFilePath);
		wcscpy(awszFileName, wsFileName.c_str());

		*phBitmap = NULL;

		CComPtr<IShellFolder> psfDesktop;
		hr = SHGetDesktopFolder(&psfDesktop); 
		if (FAILED(hr))
		{
			ERR("SHGetDesktopFolder fails");
			ERR(hr);
			break;
		}

		CComPtr<IShellFolder> psfWorkDir;
		hr = psfDesktop->ParseDisplayName(NULL, NULL, awszFilePath, NULL, &pidlWorkDir, NULL);
		if (FAILED(hr))
		{
			ERR("ParseDisplayName fails");
			ERR(hr);
			break;
		}

		hr = psfDesktop->BindToObject(pidlWorkDir, NULL, IID_IShellFolder, (LPVOID*) &psfWorkDir);
		if (FAILED(hr))
		{
			ERR("BindToObject fails");
			ERR(hr);
			break;
		}

		hr = psfWorkDir->ParseDisplayName(NULL, NULL, awszFileName, NULL, &pidlURL, NULL);
		if (FAILED(hr))
		{
			ERR("ParseDisplayName fails");
			ERR(hr);
			break;
		}

		// query IExtractImage 
		CComPtr<IExtractImage> peiURL;
		hr = psfWorkDir->GetUIObjectOf(NULL, 1, (LPCITEMIDLIST*) &pidlURL, IID_IExtractImage, NULL, (LPVOID*) &peiURL);
		if (FAILED(hr))
		{
			ERR("GetUIObjectOf fails");
			ERR(hr);
			break;
		}

		// define thumbnail properties 
		SIZE size = { THUMB_CX, THUMB_CX };
		DWORD dwPriority = 0, dwFlags = IEIFLAG_ASPECT;
		WCHAR awszImagePath[MAX_PATH];
		hr = peiURL->GetLocation(awszImagePath, MAX_PATH, &dwPriority, &size, 16, &dwFlags); 
		if (FAILED(hr))
		{
			ERR("GetLocation fails");
			ERR(hr);
			break;
		}
		//E_NOINTERFACE
		// generate thumbnail 
		hr = peiURL->Extract(phBitmap); 
		if (FAILED(hr))
		{
			ERR("Extract fails");
			ERR(hr);
			break;
		}

	} while (false);

	// free allocated structures
	if (pidlWorkDir)
		CoTaskMemFree(pidlWorkDir);
	if (pidlURL)
		CoTaskMemFree(pidlURL);

	END("CWindowsThumbnail::GetShellThumbnailImageVistaDown");
	return hr;
}

HRESULT CWindowsThumbnail::GetShellThumbnailImageVistaUp(LPCWSTR pwszFilePath, HBITMAP* phBitmap)
{
	START("CWindowsThumbnail::GetShellThumbnailImageVistaUp");
	HRESULT hr = E_FAIL;
	WCHAR awszFilePath[MAX_PATH], awszFileName[MAX_PATH];

	do
	{
		if (pwszFilePath == NULL)
		{
			break;
		}

		if (phBitmap == NULL)
		{
			break;
		}

		wcscpy(awszFilePath, pwszFilePath);

		wstring wsFileName = CStringUtil::GetFileName(pwszFilePath);
		wcscpy(awszFileName, wsFileName.c_str());

		*phBitmap = NULL;

		hr = GetShellThumbnailImageFromCache(pwszFilePath, phBitmap);
		if (SUCCEEDED(hr))
		{
			break;
		}

		hr = GetShellThumbnailImageFromProvider(pwszFilePath, phBitmap);
		if (SUCCEEDED(hr))
		{
			break;
		}

		hr = GetShellThumbnailImageFromFactory(pwszFilePath, phBitmap);
		if (SUCCEEDED(hr))
		{
			break;
		}
	} while (false);

	END("CWindowsThumbnail::GetShellThumbnailImageVistaUp");
	return hr;
}

HRESULT	CWindowsThumbnail::GetShellThumbnailImageFromProvider(LPCWSTR pwszFilePath, HBITMAP* phBitmap)
{
	START("CWindowsThumbnail::GetShellThumbnailImageFromProvider");
	HRESULT hr = E_FAIL;
	WCHAR awszFilePath[MAX_PATH], awszFileName[MAX_PATH];

	do
	{
		if (pwszFilePath == NULL)
		{
			break;
		}

		if (phBitmap == NULL)
		{
			break;
		}

		wcscpy(awszFilePath, pwszFilePath);

		wstring wsFileName = CStringUtil::GetFileName(pwszFilePath);
		wcscpy(awszFileName, wsFileName.c_str());

		*phBitmap = NULL;

		//hr = SHCreateItemFromParsingName(awszFilePath, NULL, IID_PPV_ARGS(&m_pShellItem));
		if (g_pfnSHCreateItemFromParsingName != NULL)
		{
			hr = g_pfnSHCreateItemFromParsingName(awszFilePath, NULL, IID_PPV_ARGS(&m_pShellItem));
		}
		if (SUCCEEDED(hr))
		{
			hr = m_pShellItem->BindToHandler(NULL, BHID_ThumbnailHandler, IID_PPV_ARGS(&m_pThumbProvider));
			if (SUCCEEDED(hr))
			{
				WTS_ALPHATYPE dwAlpha;
				hr = m_pThumbProvider->GetThumbnail(THUMB_CX, phBitmap, &dwAlpha);
				if (!SUCCEEDED(hr))
				{
					ERR("GetThumbnail fails");
					ERR(hr);
				}

				//m_pThumbProvider->Release();
			}
			else
			{
				ERR("BindToHandler fails");
				ERR(hr);
			}

			//m_pShellItem->Release();
		}
		else
		{
			ERR("SHCreateItemFromParsingName fails");
			ERR(hr);
		}

	} while (false);

	END("CWindowsThumbnail::GetShellThumbnailImageFromProvider");
	return hr;
}

HRESULT	CWindowsThumbnail::GetShellThumbnailImageFromCache(LPCWSTR pwszFilePath, HBITMAP* phBitmap)
{
	START("CWindowsThumbnail::GetShellThumbnailImageFromCache");
	HRESULT hr = E_FAIL;
	
	WCHAR awszFilePath[MAX_PATH], awszFileName[MAX_PATH];

	do
	{
		if (pwszFilePath == NULL)
		{
			break;
		}

		if (phBitmap == NULL)
		{
			break;
		}

		wcscpy(awszFilePath, pwszFilePath);

		wstring wsFileName = CStringUtil::GetFileName(pwszFilePath);
		wcscpy(awszFileName, wsFileName.c_str());

		*phBitmap = NULL;

		//hr = SHCreateItemFromParsingName(awszFilePath, NULL, IID_PPV_ARGS(&m_pShellItem));
		if (g_pfnSHCreateItemFromParsingName != NULL)
		{
			hr = g_pfnSHCreateItemFromParsingName(awszFilePath, NULL, IID_PPV_ARGS(&m_pShellItem));
		}
		if (SUCCEEDED(hr))
		{
			// IThumbnailCache
			hr = CoCreateInstance(CLSID_LocalThumbnailCache, NULL, CLSCTX_ALL, IID_IThumbnailCache, (LPVOID*)&m_pThumbnailCache);
			if (SUCCEEDED(hr))
			{
				WTS_CACHEFLAGS WtsCacheFlags;
				WTS_THUMBNAILID WtsThumbnailID;
				hr = m_pThumbnailCache->GetThumbnail(m_pShellItem, THUMB_CX, WTS_INCACHEONLY, &m_pSharedBitmap, &WtsCacheFlags, &WtsThumbnailID);
				if (SUCCEEDED(hr))
				{
					hr = m_pSharedBitmap->GetSharedBitmap(phBitmap);
					if (!SUCCEEDED(hr))
					{
						ERR("GetSharedBitmap fails");
						ERR(hr);
					}

					DEBG(WtsCacheFlags);
					//m_pSharedBitmap->Release();
				}
				else
				{
					ERR("GetThumbnail fails");
					ERR(hr);
				}

				//m_pThumbnailCache->Release();
			}
			else
			{
				ERR("CoCreateInstance fails");
				ERR(hr);
			}

			//m_pShellItem->Release();
		}
		else
		{
			ERR("SHCreateItemFromParsingName fails");
			ERR(hr);
		}

	} while (false);

	END("CWindowsThumbnail::GetShellThumbnailImageFromCache");
	return hr;
}

HRESULT	CWindowsThumbnail::GetShellThumbnailImageFromFactory(LPCWSTR pwszFilePath, HBITMAP* phBitmap)
{
	START("CWindowsThumbnail::GetShellThumbnailImageFromFactory");
	HRESULT hr = E_FAIL;

	WCHAR awszFilePath[MAX_PATH], awszFileName[MAX_PATH];

	do
	{
		if (pwszFilePath == NULL)
		{
			break;
		}

		if (phBitmap == NULL)
		{
			break;
		}

		wcscpy(awszFilePath, pwszFilePath);

		wstring wsFileName = CStringUtil::GetFileName(pwszFilePath);
		wcscpy(awszFileName, wsFileName.c_str());

		*phBitmap = NULL;

		// IShellItemImageFactory 
		//hr = SHCreateItemFromParsingName(awszFilePath, NULL, IID_PPV_ARGS(&m_pShellItemImageFactory));
		if (g_pfnSHCreateItemFromParsingName != NULL)
		{
			hr = g_pfnSHCreateItemFromParsingName(awszFilePath, NULL, IID_PPV_ARGS(&m_pShellItemImageFactory));
		}
		if (SUCCEEDED(hr))
		{
			SIZE sz;
			sz.cx = THUMB_CX;
			sz.cy = THUMB_CX;

			hr = m_pShellItemImageFactory->GetImage(sz, SIIGBF_THUMBNAILONLY, phBitmap);
			if (!SUCCEEDED(hr))
			{
				ERR("GetImage fails");
				ERR(hr);
			}

			//m_pShellItemImageFactory->Release();
		}
		else
		{
			ERR("SHCreateItemFromParsingName fails");
			ERR(hr);
		}

	} while (false);

	END("CWindowsThumbnail::GetShellThumbnailImageFromFactory");
	return hr;
}

HBITMAP CWindowsThumbnail::HIconToHBitmap(HICON hIcon, SIZE* pszTargetSize)
{
	ICONINFO ii = { 0 };
	if (hIcon == NULL || !GetIconInfo(hIcon, &ii) || !ii.fIcon)
	{
		return NULL;
	}

	int iWidth = 0;
	int iHeight = 0;
	if (pszTargetSize != NULL)
	{
		iWidth = pszTargetSize->cx;
		iHeight = pszTargetSize->cy;
	}
	else
	{
		if (ii.hbmColor != NULL)
		{
			BITMAP bm = { 0 };
			GetObject(ii.hbmColor, sizeof(bm), &bm);

			iWidth = bm.bmWidth;
			iHeight = bm.bmHeight;
		}
	}

	if (ii.hbmColor != NULL)
	{
		DeleteObject(ii.hbmColor);
		ii.hbmColor = NULL;
	}

	if (ii.hbmMask != NULL)
	{
		DeleteObject(ii.hbmMask);
		ii.hbmMask = NULL;
	}

	if (iWidth <= 0 || iHeight <= 0)
	{
		return NULL;
	}

	int iPixelCount = iWidth * iHeight;

	HDC hDC = GetDC(NULL);
	int* pData = NULL;
	HDC hMemDC = NULL;
	HBITMAP hBmpOld = NULL;
	bool* pOpaque = NULL;
	HBITMAP hDib = NULL;
	BOOL fSuccess = FALSE;

	do
	{
		BITMAPINFOHEADER bi = { 0 };
		bi.biSize = sizeof(BITMAPINFOHEADER);    
		bi.biWidth = iWidth;
		bi.biHeight = -iHeight;  
		bi.biPlanes = 1;    
		bi.biBitCount = 32;    
		bi.biCompression = BI_RGB;
		hDib = CreateDIBSection(hDC, (BITMAPINFO*)&bi, DIB_RGB_COLORS, (VOID**)&pData, NULL, 0);
		if (hDib == NULL)
			break;

		memset(pData, 0, iPixelCount * 4);

		hMemDC = CreateCompatibleDC(hDC);
		if (hMemDC == NULL)
			break;

		hBmpOld = (HBITMAP)SelectObject(hMemDC, hDib);
		::DrawIconEx(hMemDC, 0, 0, hIcon, iWidth, iHeight, 0, NULL, DI_MASK);

		pOpaque = new(std::nothrow) bool[iPixelCount];
		if (pOpaque == NULL)
			break;
		for (int i = 0; i < iPixelCount; ++i)
		{
			pOpaque[i] = !pData[i];
		}

		memset(pData, 0, iPixelCount * 4);
		::DrawIconEx(hMemDC, 0, 0, hIcon, iWidth, iHeight, 0, NULL, DI_NORMAL);

		BOOL fPixelHasAlpha = FALSE;
		UINT* pPixel = (UINT*)pData;
		for (int i = 0; i < iPixelCount; ++i, ++pPixel)
		{
			if ((*pPixel & 0xFF000000) != 0)
			{
				fPixelHasAlpha = TRUE;
				break;
			}
		}

		if (!fPixelHasAlpha)
		{
			pPixel = (UINT*)pData;
			for (int i = 0; i < iPixelCount; ++i, ++pPixel)
			{
				if (pOpaque[i])
				{
					*pPixel |= 0xFF000000;
				}
				else
				{
					*pPixel &= 0x00FFFFFF;
				}
			}
		}

		fSuccess = TRUE;

	} while (FALSE);


	if (pOpaque != NULL)
	{
		delete []pOpaque;
		pOpaque = NULL;
	}

	if (hMemDC != NULL)
	{
		SelectObject(hMemDC, hBmpOld);
		DeleteDC(hMemDC);
	}

	ReleaseDC(NULL, hDC);

	if (!fSuccess)
	{
		if (hDib != NULL)
		{
			DeleteObject(hDib);
			hDib = NULL;
		}
	}

	return hDib;
}

BOOL CWindowsThumbnail::GetShellIcon(LPCWSTR pwszFilePath, HBITMAP* phBitmap)
{
	BOOL fRet = FALSE;
	SHFILEINFOW shfiw;
	if (SHGetFileInfoW(m_wsFilePath.data(), FILE_ATTRIBUTE_NORMAL, &shfiw, sizeof(SHFILEINFOW), SHGFI_ICON))
	{
		SIZE szIcon;
		szIcon.cx = THUMB_CX;
		szIcon.cy = THUMB_CX;

		if (phBitmap != NULL)
		{
			*phBitmap = HIconToHBitmap(shfiw.hIcon, &szIcon);
			m_hBmp = *phBitmap;
		}

		DestroyIcon(shfiw.hIcon);

		fRet = TRUE;
	}

	return fRet;
}

BOOL CWindowsThumbnail::SetFilePath(const wstring& wsFilePath)
{
	m_wsFilePath = wsFilePath;
	return TRUE;
}

BOOL CWindowsThumbnail::GetBitmap(HBITMAP* phBitmap)
{
	HRESULT hr = E_FAIL;

	do 
	{
		Release();

		if (m_wsFilePath.empty())
		{
			break;
		}

		hr = GetShellThumbnailImageVistaDown(m_wsFilePath.data(), phBitmap);
		if (SUCCEEDED(hr))
		{
			break;
		}

		hr = GetShellThumbnailImageVistaUp(m_wsFilePath.data(), phBitmap);
		if (SUCCEEDED(hr))
		{
			break;
		}

		// If all ways can't get the thumbnail, then get the file icon
		if (GetShellIcon(m_wsFilePath.data(), phBitmap))
		{
			hr = S_OK;
		}
	} while (false);
	
	return TRUE;
}

void CWindowsThumbnail::Release()
{
	if (m_pShellItem != NULL)
	{
		m_pShellItem->Release();
		m_pShellItem = NULL;
	}
	if (m_pThumbProvider != NULL)
	{
		m_pThumbProvider->Release();
		m_pThumbProvider = NULL;
	}
	if (m_pSharedBitmap != NULL)
	{
		m_pSharedBitmap->Release();
		m_pSharedBitmap = NULL;
	}
	if (m_pThumbnailCache != NULL)
	{
		m_pThumbnailCache->Release();
		m_pThumbnailCache = NULL;
	}
	if (m_pShellItemImageFactory != NULL)
	{
		m_pShellItemImageFactory->Release();
		m_pShellItemImageFactory = NULL;
	}

	if (m_hBmp != NULL)
	{
		DeleteObject(m_hBmp);
		m_hBmp = NULL;
	}
}