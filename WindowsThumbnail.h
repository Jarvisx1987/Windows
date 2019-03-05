#pragma once

class CWindowsThumbnail
{
	wstring					m_wsFilePath;
	HBITMAP					m_hBmp;

	IShellItem*				m_pShellItem;
	IThumbnailProvider*		m_pThumbProvider;
	ISharedBitmap*			m_pSharedBitmap;
	IThumbnailCache*		m_pThumbnailCache;
	IShellItemImageFactory* m_pShellItemImageFactory;
public:
	CWindowsThumbnail(void);
	~CWindowsThumbnail(void);

public:
	HRESULT					GetShellThumbnailImageVistaDown(LPCWSTR pwszFilePath, HBITMAP* phBitmap);
	HRESULT					GetShellThumbnailImageVistaUp(LPCWSTR pwszFilePath, HBITMAP* phBitmap);

	HRESULT					GetShellThumbnailImageFromProvider(LPCWSTR pwszFilePath, HBITMAP* phBitmap);
	HRESULT					GetShellThumbnailImageFromCache(LPCWSTR pwszFilePath, HBITMAP* phBitmap);
	HRESULT					GetShellThumbnailImageFromFactory(LPCWSTR pwszFilePath, HBITMAP* phBitmap);

	HBITMAP					HIconToHBitmap(HICON hIcon, SIZE* pszTargetSize);
	BOOL					GetShellIcon(LPCWSTR pwszFilePath, HBITMAP* phBitmap);

	BOOL					SetFilePath(const wstring& wsFilePath);
	BOOL					GetBitmap(HBITMAP* phBitmap);

	void					Release();
};

