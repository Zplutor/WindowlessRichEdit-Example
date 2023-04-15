#pragma once

#include <atlbase.h>

constexpr GUID CLSID_MyOLEObject = { 0xe16f8acd, 0x5b3a, 0x4167, { 0xa4, 0x49, 0xdc, 0x57, 0xd, 0xd4, 0x44, 0x59 } };

class MyOLEObject : public IOleObject, public IViewObject {
public:
    static constexpr LONG Width = 40;
    static constexpr LONG Height = 22;

    explicit MyOLEObject(ITextServices* text_service) : text_service_(text_service) {

    }

    HRESULT QueryInterface(REFIID riid, LPVOID* ppvObj) override {

        if (!ppvObj) {
            return E_INVALIDARG;
        }

        if (riid == IID_IUnknown || riid == IID_IOleObject) {
            *ppvObj = static_cast<IOleObject*>(this);
            AddRef();
            return S_OK;
        }

        if (riid == IID_IViewObject) {
            *ppvObj = static_cast<IViewObject*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObj = nullptr;
        return E_NOINTERFACE;
    }

    ULONG AddRef() override {
        return InterlockedIncrement(&reference_count_);
    }

    ULONG Release() override {
        auto result = InterlockedDecrement(&reference_count_);
        if (result == 0) {
            delete this;
        }
        return result;
    }

    HRESULT SetClientSite(IOleClientSite* pClientSite) override {
        return E_NOTIMPL;
    }

    HRESULT GetClientSite(IOleClientSite** ppClientSite) override {
        return E_NOTIMPL;
    }

    HRESULT SetHostNames(LPCOLESTR szContainerApp, LPCOLESTR szContainerObj) override {
        return E_NOTIMPL;
    }

    HRESULT Close(DWORD dwSaveOption) override {
        return E_NOTIMPL;
    }

    HRESULT SetMoniker(DWORD dwWhichMoniker, IMoniker* pmk) override {
        return E_NOTIMPL;
    }

    HRESULT GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker** ppmk) override {
        return E_NOTIMPL;
    }

    HRESULT InitFromData(IDataObject* pDataObject, BOOL fCreation, DWORD dwReserved) override {
        return E_NOTIMPL;
    }

    HRESULT GetClipboardData(DWORD dwReserved, IDataObject** ppDataObject) override {
        return E_NOTIMPL;
    }

    HRESULT DoVerb(
        LONG iVerb,
        LPMSG lpmsg,
        IOleClientSite* pActiveSite,
        LONG lindex,
        HWND hwndParent,
        LPCRECT lprcPosRect) override {
        return E_NOTIMPL;
    }

    HRESULT EnumVerbs(IEnumOLEVERB** ppEnumOleVerb) override {
        return E_NOTIMPL;
    }

    HRESULT Update(void) override {
        return E_NOTIMPL;
    }

    HRESULT IsUpToDate(void) override {
        return E_NOTIMPL;
    }

    HRESULT GetUserClassID(CLSID* pClsid) override {
        return E_NOTIMPL;
    }

    HRESULT GetUserType(DWORD dwFormOfType, LPOLESTR* pszUserType) override {
        return E_NOTIMPL;
    }

    HRESULT SetExtent(DWORD dwDrawAspect, SIZEL* psizel) override {
        return E_NOTIMPL;
    }

    HRESULT GetExtent(DWORD dwDrawAspect, SIZEL* psizel) override {
        return E_NOTIMPL;
    }

    HRESULT Advise(IAdviseSink* pAdvSink, DWORD* pdwConnection) override {
        return E_NOTIMPL;
    }

    HRESULT Unadvise(DWORD dwConnection) override {
        return E_NOTIMPL;
    }

    HRESULT EnumAdvise(IEnumSTATDATA** ppenumAdvise) override {
        return E_NOTIMPL;
    }

    HRESULT GetMiscStatus(DWORD dwAspect, DWORD* pdwStatus) override {
        return E_NOTIMPL;
    }

    HRESULT SetColorScheme(LOGPALETTE* pLogpal) override {
        return E_NOTIMPL;
    }

    HRESULT Draw(
        DWORD dwDrawAspect,
        LONG lindex,
        void* pvAspect,
        DVTARGETDEVICE* ptd,
        HDC hdcTargetDev,
        HDC hdcDraw,
        LPCRECTL lprcBounds,
        LPCRECTL lprcWBounds,
        BOOL(*pfnContinue)(ULONG_PTR dwContinue),
        ULONG_PTR dwContinue) override {

        if (dwDrawAspect != DVASPECT_CONTENT) {
            return E_NOTIMPL;
        }

        RECT rect{};
        rect.left = lprcBounds->left;
        rect.top = lprcBounds->top;
        rect.right = lprcBounds->right;
        rect.bottom = lprcBounds->bottom;

        auto background_color = IsSelected() ? RGB(0x88, 0xaa, 0xcc) : RGB(0xaa, 0xcc, 0xee);
        HBRUSH brush = CreateSolidBrush(background_color);
        FillRect(hdcDraw, &rect, brush);
        DeleteObject(brush);

        auto old_background_mode = SetBkMode(hdcDraw, TRANSPARENT);
        TextOut(hdcDraw, rect.left + 5, rect.top, L"OLE", 3);
        SetBkMode(hdcDraw, old_background_mode);
        return S_OK;
    }

    HRESULT GetColorSet(DWORD dwDrawAspect, LONG lindex, void* pvAspect, DVTARGETDEVICE* ptd, HDC hicTargetDev, LOGPALETTE** ppColorSet) override {
        return E_NOTIMPL;
    }

    HRESULT Freeze(DWORD dwDrawAspect, LONG lindex, void* pvAspect, DWORD* pdwFreeze) override {
        return E_NOTIMPL;
    }

    HRESULT Unfreeze(DWORD dwFreeze) override {
        return E_NOTIMPL;
    }

    HRESULT SetAdvise(DWORD aspects, DWORD advf, IAdviseSink* pAdvSink) override {
        return E_NOTIMPL;
    }

    HRESULT GetAdvise(DWORD* pAspects, DWORD* pAdvf, IAdviseSink** ppAdvSink) override {
        return E_NOTIMPL;
    }

private:
    bool IsSelected() const {

        CHARRANGE select_range{};
        text_service_->TxSendMessage(EM_EXGETSEL, 0, reinterpret_cast<LPARAM>(&select_range), nullptr);

        if (select_range.cpMin == select_range.cpMax) {
            return false;
        }

        CComPtr<IRichEditOle> rich_edit_ole{};
        text_service_->TxSendMessage(EM_GETOLEINTERFACE, 0, (LPARAM)&rich_edit_ole, nullptr);

        auto object_count = rich_edit_ole->GetObjectCount();
        for (int index = 0; index < object_count; ++index) {

            REOBJECT object_info{};
            object_info.cbStruct = sizeof(object_info);
            HRESULT hresult = rich_edit_ole->GetObject(index, &object_info, REO_GETOBJ_POLEOBJ);
            if (FAILED(hresult)) {
                continue;
            }

            //Use CComPtr to auto release the OLE object.
            CComPtr<IOleObject> ole_object;
            ole_object.Attach(object_info.poleobj);

            if (ole_object.p == this) {

                if ((select_range.cpMin == 0 && select_range.cpMax == -1) ||
                    (select_range.cpMin <= object_info.cp && object_info.cp < select_range.cpMax)) {

                    return true;
                }
            }
        }

        return false;
    }

    LONG reference_count_{ 1 };
    ITextServices* text_service_{};
};