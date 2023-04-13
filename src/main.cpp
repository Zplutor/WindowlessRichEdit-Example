#include <Windows.h>
#include <atlbase.h>
#include <atlwin.h>
#include <Richedit.h>
#include <richole.h>
#include <TextServ.h>
#include <memory>
#include "my_text_host.h"
#include "my_ole_object.h"
#include "resource.h"

EXTERN_C const IID IID_ITextServices = {
    0x8d33f740,
    0xcf58,
    0x11ce,
    { 0xa8, 0x9d, 0x00, 0xaa, 0x00, 0x6c, 0xad, 0xc5 }
};

EXTERN_C const IID IID_ITextHost = {
    0xc5bdd8d0,
    0xd26e,
    0x11ce,
    { 0xa8, 0x9e, 0x00, 0xaa, 0x00, 0x6c, 0xad, 0xc5 }
};

CComPtr<MyTextHost> g_text_host;
CComPtr<ITextServices> g_text_service;

LRESULT CALLBACK WindowProcedure(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);

int WINAPI WinMain(HINSTANCE, HINSTANCE, char*, int) {

    WNDCLASSEX default_class = { 0 };
    default_class.cbSize = sizeof(default_class);
    default_class.style = CS_HREDRAW | CS_VREDRAW;
    default_class.lpfnWndProc = WindowProcedure;
    default_class.cbClsExtra = 0;
    default_class.cbWndExtra = sizeof(LONG_PTR);
    default_class.hInstance = NULL;
    default_class.hIcon = NULL;
    default_class.hCursor = LoadCursor(NULL, IDI_APPLICATION);
    default_class.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    default_class.lpszMenuName = nullptr;
    default_class.lpszClassName = L"WindowlessRichEdit";
    default_class.hIconSm = NULL;

    RegisterClassEx(&default_class);

    HWND window_handle = CreateWindowEx(
        0,
        L"WindowlessRichEdit",
        nullptr,
        WS_OVERLAPPEDWINDOW | WS_VISIBLE,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        nullptr,
        LoadMenu(nullptr, MAKEINTRESOURCE(IDR_MENU)),
        nullptr,
        nullptr
    );

    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return 0;
}

LRESULT CALLBACK WindowProcedure(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {

    switch (message) {
    case WM_CREATE: {

        g_text_host.Attach(new MyTextHost(hwnd));

        HMODULE module_handle = LoadLibrary(L"riched20.dll");

        typedef HRESULT(_stdcall* CreateTextServicesFunction)(IUnknown*, ITextHost*, IUnknown**);
        CreateTextServicesFunction create_function = reinterpret_cast<CreateTextServicesFunction>(GetProcAddress(module_handle, "CreateTextServices"));

        const IID* iid_text_service = reinterpret_cast<IID*>(GetProcAddress(module_handle, "IID_ITextServices"));

        CComPtr<IUnknown> unknown;
        create_function(nullptr, g_text_host, &unknown);
        unknown->QueryInterface(*iid_text_service, reinterpret_cast<void**>(&g_text_service));

        g_text_service->TxSetText(L"Windowless RichEdit");
        return 0;
    }

    case WM_PAINT: {

        PAINTSTRUCT paint_struct;
        HDC hdc = BeginPaint(hwnd, &paint_struct);

        RECT rect;
        GetClientRect(hwnd, &rect);

        g_text_service->TxDraw(
            DVASPECT_CONTENT,
            0,
            nullptr,
            nullptr,
            hdc,
            nullptr,
            reinterpret_cast<LPCRECTL>(&rect),
            nullptr,
            nullptr,
            nullptr,
            0,
            0
        );

        EndPaint(hwnd, &paint_struct);
        return 0;
    }

    case WM_SETFOCUS: {

        g_text_service->OnTxInPlaceActivate(nullptr);

        LRESULT result = 0;
        g_text_service->TxSendMessage(message, wParam, lParam, &result);

        return result;
    }

    case WM_KILLFOCUS: {

        g_text_service->OnTxInPlaceDeactivate();

        LRESULT result = 0;
        g_text_service->TxSendMessage(message, wParam, lParam, &result);

        return result;
    }

    case WM_KEYDOWN:
    case WM_KEYUP:
    case WM_CHAR: {

        LRESULT result = 0;
        g_text_service->TxSendMessage(message, wParam, lParam, &result);

        return result;
    }

    case WM_SETCURSOR: {
        if (LOWORD(lParam) == HTCLIENT) {

            HDC hdc = GetDC(hwnd);

            POINT position = { 0 };
            GetCursorPos(&position);
            ScreenToClient(hwnd, &position);

            RECT rect = { 0 };
            GetClientRect(hwnd, &rect);

            g_text_service->OnTxSetCursor(
                DVASPECT_CONTENT,
                0,
                nullptr,
                nullptr,
                hdc,
                nullptr,
                &rect,
                position.x,
                position.y
            );

            ReleaseDC(hwnd, hdc);
            return TRUE;
        }
        break;
    }

    case WM_MOUSEMOVE:
    case WM_LBUTTONDOWN:
    case WM_LBUTTONUP: {

        LRESULT result = 0;
        g_text_service->TxSendMessage(message, wParam, lParam, &result);
        return result;
    }

    case WM_DESTROY:
        //TextService must be released before TextHost, because the reference count of TextHost
        //is not increased when creating TextService.
        g_text_service.Release();
        g_text_host.Release();
        PostQuitMessage(0);
        return 0;

    case WM_COMMAND: {
        if (LOWORD(wParam) == ID_INSERTOLEOBJECT) {

            CComPtr<IRichEditOle> rich_edit_ole{};
            g_text_service->TxSendMessage(EM_GETOLEINTERFACE, 0, (LPARAM)&rich_edit_ole, nullptr);

            CComPtr<IOleClientSite> client_site{};
            rich_edit_ole->GetClientSite(&client_site);

            CComPtr<MyOLEObject> ole_object;
            ole_object.Attach(new MyOLEObject(g_text_service));

            REOBJECT object_info{};
            object_info.cbStruct = sizeof(object_info);
            object_info.clsid = CLSID_MyOLEObject;
            object_info.poleobj = ole_object;
            object_info.polesite = client_site;
            object_info.pstg = nullptr;
            object_info.dvaspect = DVASPECT_CONTENT;
            object_info.cp = REO_CP_SELECTION;
            object_info.dwFlags = REO_BELOWBASELINE | REO_OWNERDRAWSELECT;

            SIZEL size_in_pixels{};
            size_in_pixels.cx = MyOLEObject::Width;
            size_in_pixels.cy = MyOLEObject::Height;
            AtlPixelToHiMetric(&size_in_pixels, &object_info.sizel);

            rich_edit_ole->InsertObject(&object_info);
            return 0;
        }
        break;
    }

    default:
        break;
    }

    return CallWindowProc(DefWindowProc, hwnd, message, wParam, lParam);
}