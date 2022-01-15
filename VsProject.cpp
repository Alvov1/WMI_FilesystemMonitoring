#include <iostream>
#include <fstream>
#include <string>
#include <vector>
#include <filesystem>
#include <list>
#include <thread>

#include <comdef.h>
#include <Wbemidl.h>
#include <wincred.h>
#include <strsafe.h>

#pragma comment(lib, "wbemuuid.lib")
#pragma comment(lib, "credui.lib")
#pragma comment(lib, "comsuppw.lib")
#define _WIN32_DCOM
#define UNICODE
using namespace std;

#include "QuerySink.h"

constexpr auto configName = "config.txt";
const unsigned sleepFor = 3;

std::list<std::string> commands;

std::string doubleBackslash(std::string path) {
    for (unsigned long ind = path.find('\\'); ind != std::string::npos; ind = path.find('\\', ++(++ind)))
        path.replace(ind, 1, "\\\\");
    return path;
}
void add(std::list<std::string>& commands, const std::filesystem::path& path) {
    static const std::string beginning = "Select * From __InstanceOperationEvent Within 2 Where TargetInstance Isa ";
    const std::string Drive = "'" + path.root_name().string() + "'";
    const std::string NameAndExtension = path.filename().string();

    if (NameAndExtension.find(".") != std::string::npos) {
        /* File. */
        const std::string Path = "'" + doubleBackslash(path.parent_path().make_preferred().string()) + "\\\\'";
        const std::string Filename = "'" + path.filename().replace_extension("").string() + "'";
        const std::string Extension = "'" + path.extension().string().substr(1) + "'";
                
        const std::string comm = beginning + "'CIM_DataFile'" +
            " and TargetInstance.Drive=" + Drive +
            " and TargetInstance.Path=" + Path +
            " and TargetInstance.FileName=" + Filename +
            " and TargetInstance.Extension=" + Extension;
        commands.push_back(comm);
        return;
    } else {
        /* Directory. */
        const std::string Path = "'\\\\" + doubleBackslash(
            std::filesystem::relative(path, path.root_directory()).make_preferred().string()) + "\\\\'";

        const std::string comm1 = beginning + "'CIM_Directory'" + " and TargetInstance.Drive=" + Drive + " and TargetInstance.Path=" + Path;
        const std::string comm2 = beginning + "'CIM_DataFile'" + " and TargetInstance.Drive=" + Drive + " and TargetInstance.Path=" + Path;
        commands.push_back(comm1);
        commands.push_back(comm2);

        if(std::filesystem::exists(path))
            for (const auto& item : std::filesystem::recursive_directory_iterator(path)) 
                add(commands, item.path());
        return;
    }
}

int __cdecl main(int argc, char** argv) {
    HRESULT hres = CoInitializeEx(0, COINIT_MULTITHREADED);    
    if (FAILED(hres)) {
        cout << "Failed to initialize COM library. Error code = 0x"
            << hex << hres << endl;
        return 1;
    }
    
    hres = CoInitializeSecurity(
        nullptr,
        -1,
        nullptr,
        nullptr,
        RPC_C_AUTHN_LEVEL_DEFAULT,
        RPC_C_IMP_LEVEL_IDENTIFY,
        nullptr,
        EOAC_NONE,
        nullptr
    );

    if (FAILED(hres)) {
        cout << "Failed to initialize security. Error code = 0x"
            << hex << hres << endl;
        CoUninitialize();
        return 1;
    }
    
    IWbemLocator* pLoc = nullptr;
    
    hres = CoCreateInstance(
        CLSID_WbemLocator,
        nullptr,
        CLSCTX_INPROC_SERVER,
        IID_IWbemLocator, (LPVOID*)&pLoc);
    
    if (FAILED(hres)) {
        cout << "Failed to create IWbemLocator object."
            << " Err code = 0x"
            << hex << hres << endl;
        CoUninitialize();
        return 1;
    }
    
    IWbemServices* pSvc = nullptr;
    
    /* ----------------------------------------------------------------------------------------------------------------- */
        
    hres = pLoc->ConnectServer(
        _bstr_t(L"root\\cimv2"),
        nullptr,
        nullptr,
        nullptr,
        NULL,
        nullptr,
        nullptr,
        &pSvc
    );

    if (FAILED(hres)) {
        cout << "Could not connect. Error code = 0x"
            << hex << hres << endl;
        pLoc->Release();
        CoUninitialize();
        return 1;
    }
    cout << "Connected to ROOT\\CIMV2 WMI namespace" << endl;

    /* ----------------------------------------------------------------------------------------------------------------- */

    hres = CoSetProxyBlanket(
        pSvc,
        RPC_C_AUTHN_WINNT,
        RPC_C_AUTHZ_NONE,
        nullptr,
        RPC_C_AUTHN_LEVEL_CALL,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        nullptr,
        EOAC_NONE
    );
    if (FAILED(hres)) {
        cout << "Could not set proxy blanket. Error code = 0x"
            << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;
    }

    /* ----------------------------------------------------------------------------------------------------------------- */
    
    IUnsecuredApartment* pUnsecApp = nullptr;
    hres = CoCreateInstance(CLSID_UnsecuredApartment, nullptr, 
        CLSCTX_LOCAL_SERVER, IID_IUnsecuredApartment, (void**)&pUnsecApp);

    auto* pSink = new QuerySink;
    pSink->AddRef();

    IUnknown* pStubUnk = nullptr;
    pUnsecApp->CreateObjectStub(pSink, &pStubUnk);
    IWbemObjectSink* pStubSink = nullptr;
    pStubUnk->QueryInterface(IID_IWbemObjectSink, (void**)&pStubSink);

    IUnsecuredApartment* pUnsecAppCreate = nullptr;
    hres = CoCreateInstance(CLSID_UnsecuredApartment, nullptr, 
        CLSCTX_LOCAL_SERVER, IID_IUnsecuredApartment, (void**)&pUnsecAppCreate);
    
    /* ---------------------------------------- Reading configuration file --------------------------------------------- */

    std::ifstream config(configName);
    if (config.fail())
        throw std::runtime_error("Unable to open the config file.");
    std::string line;

    size_t pathsNumber;
    try {
        std::getline(config, line);
        pathsNumber = std::stoi(line);
    }
    catch (const std::invalid_argument& e) {
        throw std::runtime_error("Please enter the number of paths in the first line of the config.");
    }

    std::cout << "Monitoring " << pathsNumber << " files or directories:" << std::endl;
    for (auto i = 0; i < pathsNumber; ++i) {
        std::getline(config, line);

        try {
            const std::filesystem::path path(line);
            add(commands, path);
            if (std::filesystem::is_regular_file(path))
                add(commands, path.parent_path());
            std::cout << line << std::endl;
        }
        catch (const std::invalid_argument& e) {
            std::cout << "Invalid path on line " << i + 1 << ": " << line << std::endl;
        }
    }
    config.close();

    /* -------------------------------------------- Looking for events ------------------------------------------------- */

    while (true) {
        for (const auto& item : commands) {
            hres = pSvc->ExecNotificationQueryAsync(
                _bstr_t("WQL"),
                _bstr_t(item.c_str()),
                WBEM_FLAG_SEND_STATUS,
                NULL,
                pStubSink);
        }
        std::this_thread::sleep_for(std::chrono::seconds(sleepFor));
        hres = pSvc->CancelAsyncCall(pStubSink);
    }

    /* ----------------------------------------------------------------------------------------------------------------- */

    pSvc->Release();
    pLoc->Release();
    pUnsecApp->Release();
    pStubUnk->Release();
    pSink->Release();
    pStubSink->Release();
    CoUninitialize();
    return 0;
}
