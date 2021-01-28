#include "stdafx.h"


extern "C" {

	// This function displays an error message
	void DisplayError(DWORD dwErr, LPCTSTR sTitle) {
		TCHAR sBuf[512];
		DWORD dwBufSize = sizeof(sBuf) / sizeof(sBuf[0]);
		if (!FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, dwErr, 0, sBuf, dwBufSize, NULL)) {
			wsprintf(sBuf, _T("err 0x%08X"), dwErr);
		}
		MessageBox(NULL, sBuf, sTitle, MB_OK);
	}

	static LPCTSTR MyEventName = _T("{78FF3DDE-04DF-44eb-9FFB-EA6468418DCD}");
	static LPCTSTR ConcurrentInstallProperty = _T("CONCURRENTINSTALL");

	UINT __stdcall ForceUniqueInstall(MSIHANDLE hInstall)
	{
		// Clear last error
		SetLastError(0);
		HANDLE hEvent = CreateEvent(NULL, FALSE, FALSE, MyEventName);
		if ((hEvent != NULL) && (GetLastError() == ERROR_ALREADY_EXISTS)) {
			// Close the handle as we don't need it anymore
			CloseHandle(hEvent);

			// And set the property
			MsiSetProperty(hInstall, ConcurrentInstallProperty, _T("True"));
		}
		return ERROR_SUCCESS;
	}

	typedef struct _ProductDetection {
		LPCTSTR ProductCode;
		LPCTSTR ProductProperty;
	} ProductDetection;

	static ProductDetection ProductsToDetect[] = {
		{	_T("{782F78A8-9388-43FC-ACF9-1F345591AFEF}"),	_T("DAISYWORD2003")	},
		{	_T("{CE315649-024F-4BE5-B6CA-05602436D4A1}"),	_T("DAISYWORD2007")	},
		{	_T("{CFD34E50-1596-42A0-BD51-1302EC5D7F5B}"),	_T("DAISYWORDXP")	},
		{	_T("{91490409-6000-11D3-8CFE-0150048383C9}"),	_T("PIAWORD2003")	},
		{	_T("{50120000-1105-0000-0000-0000000FF1CE}"),	_T("PIAWORD2007")	}
	};

	static LPCTSTR WordVersionProperty = L"WORDVERSION";

	// This function will check for the presence of old versions of the converter
	UINT __stdcall DetectPreviousConverters(MSIHANDLE hInstall) {
		int productsCount = sizeof(ProductsToDetect) / sizeof(ProductsToDetect[0]);
		for (int i=0;i<productsCount;i++) {
			TCHAR sBuf[512];
			DWORD dwBufSize = sizeof(sBuf) / sizeof(sBuf[0]);
			int nRet = MsiGetProductInfo(ProductsToDetect[i].ProductCode, INSTALLPROPERTY_PRODUCTNAME, sBuf, &dwBufSize);
			switch (nRet) {
			case ERROR_SUCCESS:
				// Product is installed
				OutputDebugString(_T("Product detected : "));
				OutputDebugString(ProductsToDetect[i].ProductProperty);
				OutputDebugString(_T("\n"));
				nRet = MsiSetProperty(hInstall, ProductsToDetect[i].ProductProperty, _T("True"));
				if (nRet != ERROR_SUCCESS) {
					DisplayError(nRet, _T("MsiSetProperty"));
				}
				break;
			case ERROR_UNKNOWN_PRODUCT:
				// Product is not installed
				OutputDebugString(_T("Product NOT detected : "));
				OutputDebugString(ProductsToDetect[i].ProductProperty);
				OutputDebugString(_T("\n"));
				break;
			default:
				DisplayError(nRet, _T("MsiGetProductInfo"));
				break;
			}
		}
		return ERROR_SUCCESS;
	}


	UINT __stdcall GetWordVersion(MSIHANDLE hInstall) {
		HKEY hKey;
		UINT nRet = RegOpenKeyEx(HKEY_CLASSES_ROOT, _T("Word.Application\\Curver"), 0, KEY_READ, &hKey);
		if (nRet == ERROR_SUCCESS) {
			TCHAR sBuf[512];
			DWORD dwBufSize = sizeof(sBuf) / sizeof(sBuf[0]);
			DWORD dwType;
			nRet = RegQueryValueEx(hKey, NULL, NULL, &dwType, (LPBYTE)sBuf, &dwBufSize);
			RegCloseKey(hKey);

			if (nRet == ERROR_SUCCESS) {
				if (!_tcsnicmp(_T("Word.Application."), sBuf, 17)) {
					// Fine
					OutputDebugString(_T("Word version detected : "));
					OutputDebugString(sBuf + 17);
					OutputDebugString(_T("\n"));
					nRet = MsiSetProperty(hInstall, WordVersionProperty, sBuf + 17);
					if (nRet != ERROR_SUCCESS) {
						DisplayError(nRet, _T("MsiSetProperty"));
					}
	
					if(!wcscmp(_T("12"),sBuf+17))
					{
						nRet = MsiSetProperty(hInstall, ProductsToDetect[3].ProductProperty, _T("True"));
					}
					else if(!wcscmp(_T("11"),sBuf+17))
					{
						nRet = MsiSetProperty(hInstall, ProductsToDetect[4].ProductProperty, _T("True"));
					}
					else 
					{
						nRet = MsiSetProperty(hInstall, ProductsToDetect[3].ProductProperty, _T("True"));
						nRet = MsiSetProperty(hInstall, ProductsToDetect[4].ProductProperty, _T("True"));
					}

					return ERROR_SUCCESS;
				}
			}
		}
		// In any case except full detection succeeded
		OutputDebugString(_T("No word version detected"));
		MsiSetProperty(hInstall, WordVersionProperty, L"None");
		return ERROR_SUCCESS;
	}

	static LPCTSTR PropertiesTpDump[] = {
		_T("GUDULEINSTALLED"),
		_T("DAISYWORD2003"),
		_T("DAISYWORD2007"),
		_T("DAISYWORDXP"),
		_T("PIAWORD2003"),
		_T("PIAWORD2007")
	};

	UINT __stdcall DumpProperties(MSIHANDLE hInstall) {
		int productsCount = sizeof(PropertiesTpDump) / sizeof(PropertiesTpDump[0]);
		for (int i=0;i<productsCount;i++) {
			TCHAR sBuf[512];
			DWORD dwBufSize = sizeof(sBuf) / sizeof(sBuf[0]);
			UINT nRet = MsiGetProperty(hInstall, PropertiesTpDump[i], sBuf, &dwBufSize);
			if (nRet == ERROR_SUCCESS) {
				OutputDebugString(PropertiesTpDump[i]);
				OutputDebugString(_T(" : "));
				OutputDebugString(sBuf);
				OutputDebugString(_T("\n"));
			} else {
				DisplayError(nRet, _T("MsiGetProperty"));
			}
		}
		return ERROR_SUCCESS;
	}

	UINT __stdcall LaunchReadme(MSIHANDLE hInstall) {
		TCHAR sBuf[512];
		DWORD dwBufSize = sizeof(sBuf) / sizeof(sBuf[0]);
		UINT nRet = MsiGetProperty(hInstall, _T("TARGETDIR"), sBuf, &dwBufSize);
		if (nRet == ERROR_SUCCESS) {
			ShellExecute(NULL, _T("Open"), _T("Help.chm"), NULL, sBuf, SW_SHOW);
		} else {
			DisplayError(nRet, _T("MsiGetProperty"));
		}
		return ERROR_SUCCESS;
	}
} // extern "C"