#pragma once
#include <iostream>
#include <list>
#include <string>
#include <wbemidl.h>
#pragma comment(lib, "wbemuuid.lib")

extern std::list<std::string> commands;
extern void add(std::list<std::string>& commands, const std::filesystem::path& path);

class QuerySink : public IWbemObjectSink {
    LONG m_lRef;
    bool bDone;

public:
    QuerySink() { m_lRef = 0; }
    ~QuerySink() { bDone = TRUE; }

    virtual ULONG STDMETHODCALLTYPE AddRef();
    virtual ULONG STDMETHODCALLTYPE Release();
    virtual HRESULT STDMETHODCALLTYPE
        QueryInterface(REFIID riid, void** ppv);

    virtual HRESULT STDMETHODCALLTYPE Indicate(
        LONG lObjectCount,
        IWbemClassObject __RPC_FAR* __RPC_FAR* apObjArray
    );

    virtual HRESULT STDMETHODCALLTYPE SetStatus(
        /* [in] */ LONG lFlags,
        /* [in] */ HRESULT hResult,
        /* [in] */ BSTR strParam,
        /* [in] */ IWbemClassObject __RPC_FAR* pObjParam
    );
};


ULONG QuerySink::AddRef() {
    return InterlockedIncrement(&m_lRef);
}

ULONG QuerySink::Release() {
    LONG lRef = InterlockedDecrement(&m_lRef);
    if (lRef == 0)
        delete this;
    return lRef;
}

HRESULT QuerySink::QueryInterface(REFIID riid, void** ppv) {
    if (riid == IID_IUnknown || riid == IID_IWbemObjectSink)
    {
        *ppv = (IWbemObjectSink*)this;
        AddRef();
        return WBEM_S_NO_ERROR;
    }
    else return E_NOINTERFACE;
}


HRESULT QuerySink::Indicate(long lObjCount, IWbemClassObject** pArray) {
    for (long i = 0; i < lObjCount; i++)
    {
        IWbemClassObject* pObj = pArray[i];
        _variant_t classVariant;
        _variant_t targetVariant;
        _variant_t captionVariant;

        auto hres = pObj->Get(_bstr_t(L"__Class"), 0, &classVariant, nullptr, nullptr);
        hres = pObj->Get(_bstr_t(L"TargetInstance"), 0, &targetVariant, nullptr, nullptr);
        hres = ((IUnknown*)targetVariant)->QueryInterface(IID_IWbemClassObject, reinterpret_cast<void**>(&pObj));
        hres = pObj->Get(L"Caption", 0, &captionVariant, nullptr, nullptr);

        if (SUCCEEDED(hres)) {
            const std::wstring classOrg(classVariant.bstrVal);

            if (classOrg == L"__InstanceDeletionEvent")
                std::wcout << L" - Deleted '" << captionVariant.bstrVal << "'." << std::endl;

            if (classOrg == L"__InstanceCreationEvent") {
                std::wcout << L" + Created '" << captionVariant.bstrVal << "'." << std::endl;
                
                try {
                    const std::filesystem::path path(captionVariant.bstrVal);
                    add(commands, path);
                }
                catch (const std::invalid_argument& e) {
                    std::cout << "Error: adding to the paths storage." << std::endl;
                }
            }

            if (classOrg == L"__InstanceModificationEvent")
                std::wcout << L" # Modified '" << captionVariant.bstrVal << "'." << std::endl;
        }

        VariantClear(&captionVariant);
    }

    return WBEM_S_NO_ERROR;
}

HRESULT QuerySink::SetStatus(
    /* [in] */ LONG lFlags,
    /* [in] */ HRESULT hResult,
    /* [in] */ BSTR strParam,
    /* [in] */ IWbemClassObject __RPC_FAR* pObjParam) {
    return WBEM_S_NO_ERROR;
}
