/////////////////////////////////////////////////////////////////////////////
// Copyright © 2013 Robert McNeel & Associates. All rights reserved.
// Rhinoceros is a registered trademark of Robert McNeel & Assoicates.
//
// THIS SOFTWARE IS PROVIDED "AS IS" WITHOUT EXPRESS OR IMPLIED WARRANTY.
// ALL IMPLIED WARRANTIES OF FITNESS FOR ANY PARTICULAR PURPOSE AND OF
// MERCHANTABILITY ARE HEREBY DISCLAIMED.
//

#include "stdafx.h"

bool FileExists(const wchar_t* pszFileName)
{
  bool rc = false;
  if (0 == pszFileName || 0 == pszFileName[0])
    return rc;

  WIN32_FIND_DATA FindFileData;
  memset(&FindFileData, 0, sizeof(WIN32_FIND_DATA));

  HANDLE hFind = FindFirstFile(pszFileName, &FindFileData);
  rc = (hFind != INVALID_HANDLE_VALUE);
  if (rc)
    FindClose(hFind);

  return rc;
}

bool DoRegister(const wchar_t* pszFlag)
{
  bool rc = true;
  if (0 == pszFlag || 0 == pszFlag[0])
    return rc;

  if (0 == _wcsicmp(pszFlag, L"/register"))
    rc = true;
  else if (0 == _wcsicmp(pszFlag, L"-register"))
    rc = true;
  else if (0 == _wcsicmp(pszFlag, L"/r"))
    rc = true;
  else if (0 == _wcsicmp(pszFlag, L"-r"))
    rc = true;
  else if (0 == _wcsicmp(pszFlag, L"/unregister"))
    rc = false;
  else if (0 == _wcsicmp(pszFlag, L"-unregister"))
    rc = false;
  else if (0 == _wcsicmp(pszFlag, L"/u"))
    rc = false;
  else if (0 == _wcsicmp(pszFlag, L"-u"))
    rc = false;

  return rc;
}

int APIENTRY wWinMain(HINSTANCE, HINSTANCE, LPWSTR, int)
{
  const wchar_t* pszAppName = L"RegTypeLib";

  int num_args = 0;
  wchar_t** pszArgs = CommandLineToArgvW(GetCommandLineW(), &num_args);
  if (0 == pszArgs || num_args < 2 || num_args > 3)
  {
    const wchar_t* pszMessage =
      L"Type Library Registration Utility\n"
      L"Copyright © 2013, Robert McNeel & Associates\n\n"
      L"Usage: RegTypeLib tlbname [/R] [/U]\n\n"
      L"/R - Registers the specified type library.\n"
      L"/U - Unregisters the specified type library.";
    MessageBox(0, pszMessage, pszAppName, MB_OK);
    return -1;
  }

  wchar_t szFileName[_MAX_PATH];
  wmemset(szFileName, 0, _MAX_PATH);
  StringCchCopy(szFileName, _MAX_PATH, pszArgs[1]);

  bool bRegister = (num_args == 3) ? DoRegister(pszArgs[2]) : true;

  LocalFree(pszArgs); // Don't leak...
  pszArgs = 0;

  if (!FileExists(szFileName))
  {
    MessageBox(0, L"File not found.", pszAppName, MB_ICONHAND);
    return -1;
  }

  CoInitialize(0);

  ITypeLib* pTypeLib = 0;
  HRESULT hResult = LoadTypeLib(szFileName, &pTypeLib);
  if (FAILED(hResult))
  {
    MessageBox(0, L"Unable to load type library.", pszAppName, MB_ICONHAND);
    CoUninitialize();
    return hResult;
  }

  if (bRegister)
  {
    hResult = RegisterTypeLib(pTypeLib, szFileName, 0);
    if (FAILED(hResult))
      MessageBox(0, L"Unable to register type library.", pszAppName, MB_ICONHAND);
    else
      MessageBox(0, L"Type Library registration successful.", pszAppName, MB_ICONINFORMATION);

    pTypeLib->Release();
    CoUninitialize();
    return hResult;
  }

  TLIBATTR* pLibAttr = 0;
  hResult = pTypeLib->GetLibAttr(&pLibAttr);
  if (FAILED(hResult))
  {
    MessageBox(0, L"Unable to get type library attributes.", pszAppName, MB_ICONHAND);
    pTypeLib->Release();
    CoUninitialize();
    return hResult;
  }

  hResult = UnRegisterTypeLib(pLibAttr->guid, pLibAttr->wMajorVerNum, pLibAttr->wMinorVerNum, pLibAttr->lcid, pLibAttr->syskind);
  pTypeLib->ReleaseTLibAttr(pLibAttr);
  pTypeLib->Release();

  if (FAILED(hResult) && 0x8002801C != hResult)
    MessageBox(0, L"Unable to unregister type library.", pszAppName, MB_ICONINFORMATION);
  else
    MessageBox(0, L"Unregistration of type library successful.", pszAppName, MB_ICONINFORMATION);

  CoUninitialize();

  return hResult;
}
