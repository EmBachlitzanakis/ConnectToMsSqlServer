// msolebdb/main.h
#pragma once

#include <windows.h>
//#include <oledb.h>
#include <comdef.h>
#include <MSOLEDBSQL.h>
#include <string>
#include <cstdio>

// MSOLEDBSQL CLSID definition
#ifndef MSOLEDBSQL_CLSID_DEFINED
#define MSOLEDBSQL_CLSID_DEFINED
// {5A23DE84-1D7B-4A16-8DED-B29C09CB648D}
DEFINE_GUID(MSOLEDBSQL_CLSID,
    0x5a23de84, 0x1d7b, 0x4a16, 0x8d, 0xed, 0xb2, 0x9c, 0x09, 0xcb, 0x64, 0x8d);
#endif

namespace msolebdb {

    inline HRESULT InitializeAndEstablishConnection(
        IDBInitialize*& pIDBInitialize,
        const std::wstring& server,
        const std::wstring& database,
        const std::wstring& uid,
        const std::wstring& pwd,
        bool trustedConnection = false,
        bool trustServerCertificate = true)
    {
        HRESULT hr = S_OK;
        IDBProperties* pIDBProperties = nullptr;

        // Build connection string
        std::wstring connectionString = L"Provider=MSOLEDBSQL;";
        connectionString += L"Server=" + server + L";";
        connectionString += L"Database=" + database + L";";
        connectionString += L"UID=" + uid + L";";
        connectionString += L"PWD=" + pwd + L";";
        connectionString += L"Trusted_Connection=no";
        connectionString += L"TrustServerCertificate=yes;";
       // connectionString += trustedConnection ? L"yes;" : L"no;";
       // if (trustServerCertificate)
      //      connectionString += L"TrustServerCertificate=yes;";

        // Create OLE DB provider instance
        hr = CoCreateInstance(MSOLEDBSQL_CLSID, nullptr, CLSCTX_INPROC_SERVER, IID_IDBInitialize, (void**)&pIDBInitialize);
        if (FAILED(hr)) {
            std::printf("Failed to obtain access to the OLE DB Driver.\n");
            return hr;
        }

        // Set connection properties
        DBPROP initProp = {};
        VariantInit(&initProp.vValue);
        initProp.dwPropertyID = DBPROP_INIT_PROVIDERSTRING;
        initProp.vValue.vt = VT_BSTR;
        initProp.vValue.bstrVal = SysAllocString(connectionString.c_str());
        initProp.dwOptions = DBPROPOPTIONS_REQUIRED;
        initProp.colid = DB_NULLID;

        DBPROPSET propSet = {};
        propSet.guidPropertySet = DBPROPSET_DBINIT;
        propSet.cProperties = 1;
        propSet.rgProperties = &initProp;

        hr = pIDBInitialize->QueryInterface(IID_IDBProperties, (void**)&pIDBProperties);
        if (FAILED(hr)) {
            std::printf("Failed to obtain an IDBProperties interface.\n");
            SysFreeString(initProp.vValue.bstrVal);
            pIDBInitialize->Release();
            pIDBInitialize = nullptr;
            return hr;
        }

        hr = pIDBProperties->SetProperties(1, &propSet);
        if (FAILED(hr)) {
            std::printf("Failed to set initialization properties.\n");
            SysFreeString(initProp.vValue.bstrVal);
            pIDBProperties->Release();
            pIDBProperties = nullptr;
            pIDBInitialize->Release();
            pIDBInitialize = nullptr;
            return hr;
        }

        // Establish the connection
        hr = pIDBInitialize->Initialize();
        if (FAILED(hr)) {
            std::printf("Failed to establish connection with the server.\n");
            _com_error err(hr);
            std::wprintf(L"Error: 0x%08X - %s\n", hr, err.ErrorMessage());
            SysFreeString(initProp.vValue.bstrVal);
            pIDBProperties->Release();
            pIDBProperties = nullptr;
            pIDBInitialize->Release();
            pIDBInitialize = nullptr;
            return hr;
        }

        // Cleanup
        SysFreeString(initProp.vValue.bstrVal);
        if (pIDBProperties) {
            pIDBProperties->Release();
            pIDBProperties = nullptr;
        }

        return hr;
    }

} // namespace msolebdb
