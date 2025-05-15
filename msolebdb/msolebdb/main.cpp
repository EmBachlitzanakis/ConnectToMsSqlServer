#include "msoledbsql.h"
#include <stdio.h>
#include <comdef.h>

// MSOLEDBSQL CLSID definition
DEFINE_GUID(MSOLEDBSQL_CLSID,
    0x5a23de84, 0x1d7b, 0x4a16, 0x8d, 0xed, 0xb2, 0x9c, 0x09, 0xcb, 0x64, 0x8d);

HRESULT InitializeAndEstablishConnection(IDBInitialize*& pIDBInitialize);

int main() {
    HRESULT hr = CoInitialize(nullptr);
    if (FAILED(hr)) {
        printf("COM initialization failed.\n");
        return 1;
    }

    IDBInitialize* pIDBInitialize = nullptr;
    hr = InitializeAndEstablishConnection(pIDBInitialize);
    if (FAILED(hr)) {
        printf("Failed to establish connection.\n");
        CoUninitialize();
        return 1;
    }

    printf("Connection to the database was established successfully.\n");

    // Cleanup
    if (pIDBInitialize) {
        pIDBInitialize->Uninitialize();
        pIDBInitialize->Release();
        pIDBInitialize = nullptr;
    }
    CoUninitialize();
    return 0;
}

HRESULT InitializeAndEstablishConnection(IDBInitialize*& pIDBInitialize) {
    HRESULT hr = S_OK;
    IDBProperties* pIDBProperties = nullptr;

    // Connection string
    const wchar_t* connectionString =
        L"Provider=MSOLEDBSQL;"
        L"Server=;" // here you write your server name 
		L"Database=;"// here you write your database name
		L"UID=;" // here you write your user name
		L"PWD=;"// here you write your password
        L"Trusted_Connection=no;"
        L"TrustServerCertificate=yes;";

    // Create OLE DB provider instance
    hr = CoCreateInstance(MSOLEDBSQL_CLSID, nullptr, CLSCTX_INPROC_SERVER, IID_IDBInitialize, (void**)&pIDBInitialize);
    if (FAILED(hr)) {
        printf("Failed to obtain access to the OLE DB Driver.\n");
        return hr;
    }

    // Set connection properties
    DBPROP initProp = {};
    VariantInit(&initProp.vValue);
    initProp.dwPropertyID = DBPROP_INIT_PROVIDERSTRING;
    initProp.vValue.vt = VT_BSTR;
    initProp.vValue.bstrVal = SysAllocString(connectionString);
    initProp.dwOptions = DBPROPOPTIONS_REQUIRED;
    initProp.colid = DB_NULLID;

    DBPROPSET propSet = {};
    propSet.guidPropertySet = DBPROPSET_DBINIT;
    propSet.cProperties = 1;
    propSet.rgProperties = &initProp;

    hr = pIDBInitialize->QueryInterface(IID_IDBProperties, (void**)&pIDBProperties);
    if (FAILED(hr)) {
        printf("Failed to obtain an IDBProperties interface.\n");
        SysFreeString(initProp.vValue.bstrVal);
        pIDBInitialize->Release();
        pIDBInitialize = nullptr;
        return hr;
    }

    hr = pIDBProperties->SetProperties(1, &propSet);
    if (FAILED(hr)) {
        printf("Failed to set initialization properties.\n");
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
        printf("Failed to establish connection with the server.\n");
        _com_error err(hr);
        wprintf(L"Error: 0x%08X - %s\n", hr, err.ErrorMessage());
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