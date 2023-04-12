#pragma once

#include <Windows.h>
#include <Richedit.h>
#include <TextServ.h>
#include <memory>

class MyTextHost : public ITextHost {
public:
    MyTextHost(HWND hwnd) : hwnd_(hwnd) { }

    HRESULT QueryInterface(REFIID riid, void** ppvObject) override {

        if (ppvObject == nullptr) {
            return E_POINTER;
        }

        if ((riid == IID_IUnknown) || (riid == IID_ITextHost)) {
            *ppvObject = this;
            return S_OK;
        }

        *ppvObject = nullptr;
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

    HDC TxGetDC() override {
        return GetDC(hwnd_);
    }

    INT TxReleaseDC(HDC hdc) override {
        return ReleaseDC(hwnd_, hdc);
    }

    BOOL TxShowScrollBar(INT fnBar, BOOL fShow) override {
        return FALSE;
    }

    BOOL TxEnableScrollBar(INT fuSBFlags, INT fuArrowflags) override {
        return FALSE;
    }

    BOOL TxSetScrollRange(INT fnBar, LONG nMinPos, INT nMaxPos, BOOL fRedraw) override {
        return FALSE;
    }

    BOOL TxSetScrollPos(INT fnBar, INT nPos, BOOL fRedraw) override {
        return FALSE;
    }

    void TxInvalidateRect(LPCRECT prc, BOOL fMode) override {
        InvalidateRect(hwnd_, prc, fMode);
    }

    void TxViewChange(BOOL fUpdate) override {

    }

    BOOL TxCreateCaret(HBITMAP hbmp, INT xWidth, INT yHeight) override {
        return CreateCaret(hwnd_, hbmp, xWidth, yHeight);
    }

    BOOL TxShowCaret(BOOL fShow) override {
        if (fShow) {
            return ShowCaret(hwnd_);
        }
        else {
            return HideCaret(hwnd_);
        }
    }

    BOOL TxSetCaretPos(INT x, INT y) override {
        return SetCaretPos(x, y);
    }

    BOOL TxSetTimer(UINT idTimer, UINT uTimeout) override {
        return FALSE;
    }

    void TxKillTimer(UINT idTimer) override {

    }

    void TxScrollWindowEx(INT dx, INT dy, LPCRECT lprcScroll, LPCRECT lprcClip, HRGN hrgnUpdate, LPRECT lprcUpdate, UINT fuScroll) override {

    }

    void TxSetCapture(BOOL fCapture) override {
        if (fCapture) {
            SetCapture(hwnd_);
        }
        else {
            ReleaseCapture();
        }
    }

    void TxSetFocus() override {

    }

    void TxSetCursor(HCURSOR hcur, BOOL fText) override {
        SetCursor(hcur);
    }

    BOOL TxScreenToClient(LPPOINT lppt) override {
        return FALSE;
    }

    BOOL TxClientToScreen(LPPOINT lppt) override {
        return FALSE;
    }

    HRESULT TxActivate(LONG* plOldState) override {
        return E_NOTIMPL;
    }

    HRESULT TxDeactivate(LONG lNewState) override {
        return E_NOTIMPL;
    }

    HRESULT TxGetClientRect(LPRECT prc) override {
        GetClientRect(hwnd_, prc);
        return S_OK;
    }

    HRESULT TxGetViewInset(LPRECT prc) override {
        return E_NOTIMPL;
    }

    HRESULT TxGetCharFormat(const CHARFORMATW** ppCF) override {

        if (char_format_ == nullptr) {
            char_format_ = std::make_unique<CHARFORMATW>();
            char_format_->cbSize = sizeof(CHARFORMATW);
        }

        *ppCF = char_format_.get();
        return S_OK;
    }

    HRESULT TxGetParaFormat(const PARAFORMAT** ppPF) override {

        if (para_format_ == nullptr) {
            para_format_ = std::make_unique<PARAFORMAT>();
            para_format_->cbSize = sizeof PARAFORMAT;
        }

        *ppPF = para_format_.get();
        return S_OK;
    }

    COLORREF TxGetSysColor(int nIndex) override {
        return GetSysColor(nIndex);
    }

    HRESULT TxGetBackStyle(TXTBACKSTYLE* pstyle) override {
        return E_NOTIMPL;
    }

    HRESULT TxGetMaxLength(DWORD* plength) override {
        return E_NOTIMPL;
    }

    HRESULT TxGetScrollBars(DWORD* pdwScrollBar) override {
        return E_NOTIMPL;
    }

    HRESULT TxGetPasswordChar(_Out_ TCHAR* pch) override {
        return E_NOTIMPL;
    }

    HRESULT TxGetAcceleratorPos(LONG* pcp) override {
        return E_NOTIMPL;
    }

    HRESULT TxGetExtent(LPSIZEL lpExtent) override {
        return E_NOTIMPL;
    }

    HRESULT OnTxCharFormatChange(const CHARFORMATW* pCF) override {
        return E_NOTIMPL;
    }

    HRESULT OnTxParaFormatChange(const PARAFORMAT* pPF) override {
        return E_NOTIMPL;
    }

    HRESULT TxGetPropertyBits(DWORD dwMask, DWORD* pdwBits) override {
        *pdwBits = 0;
        return S_OK;
    }

    HRESULT TxNotify(DWORD iNotify, void* pv) override {
        return E_NOTIMPL;
    }

    HIMC TxImmGetContext() override {
        return nullptr;
    }

    void TxImmReleaseContext(HIMC himc) override {

    }

    HRESULT TxGetSelectionBarWidth(LONG* lSelBarWidth) override {
        *lSelBarWidth = 0;
        return S_OK;
    }

private:
    LONG reference_count_{ 1 };
    HWND hwnd_{};
    std::unique_ptr<CHARFORMATW> char_format_;
    std::unique_ptr<PARAFORMAT> para_format_;
};
