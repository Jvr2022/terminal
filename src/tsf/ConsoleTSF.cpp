// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

#include "precomp.h"
#include "TfConvArea.h"
#include "TfEditSession.h"

/* 626761ad-78d2-44d2-be8b-752cf122acec */
const GUID GUID_APPLICATION = { 0x626761ad, 0x78d2, 0x44d2, { 0xbe, 0x8b, 0x75, 0x2c, 0xf1, 0x22, 0xac, 0xec } };

//+---------------------------------------------------------------------------
//
// CConsoleTSF::Initialize
//
//----------------------------------------------------------------------------

CConsoleTSF::CConsoleTSF(HWND hwndConsole, GetSuggestionWindowPos pfnPosition, GetTextBoxAreaPos pfnTextArea) :
    _hwndConsole(hwndConsole),
    _pfnPosition(pfnPosition),
    _pfnTextArea(pfnTextArea)
{
    _spITfThreadMgr = wil::CoCreateInstance<ITfThreadMgrEx, ITfThreadMgrEx>(CLSCTX_ALL);

    THROW_IF_FAILED(_spITfThreadMgr->ActivateEx(&_tid, TF_TMAE_CONSOLE));

    // Create Cicero document manager and input context.

    THROW_IF_FAILED(_spITfThreadMgr->CreateDocumentMgr(&_spITfDocumentMgr));

    TfEditCookie ecTmp;
    THROW_IF_FAILED(_spITfDocumentMgr->CreateContext(_tid, 0, static_cast<ITfContextOwnerCompositionSink*>(this), &_spITfInputContext, &ecTmp));

    // Set the context owner before attaching the context to the doc.
    const auto spSrcIC = _spITfInputContext.query<ITfSource>();
    THROW_IF_FAILED(spSrcIC->AdviseSink(IID_ITfContextOwner, static_cast<ITfContextOwner*>(this), &_dwContextOwnerCookie));
    THROW_IF_FAILED(_spITfDocumentMgr->Push(_spITfInputContext.get()));

    // Collect the active keyboard layout info.
    if (const auto spITfProfilesMgr = wil::CoCreateInstanceNoThrow<ITfInputProcessorProfileMgr>(CLSID_TF_InputProcessorProfiles, CLSCTX_ALL))
    {
        TF_INPUTPROCESSORPROFILE ipp;
        if (SUCCEEDED(spITfProfilesMgr->GetActiveProfile(GUID_TFCAT_TIP_KEYBOARD, &ipp)))
        {
            std::ignore = OnActivated(ipp.dwProfileType, ipp.langid, ipp.clsid, ipp.catid, ipp.guidProfile, ipp.hkl, ipp.dwFlags);
        }
    }

    // Setup some useful Cicero event sinks and callbacks.
    // _spITfThreadMgr && _spITfInputContext must be non-null for checks above to have succeeded, so
    // we're not going to check them again here. query will A/V if they're null.
    const auto spSrcTIM = _spITfThreadMgr.query<ITfSource>();
    const auto spSrcICS = _spITfInputContext.query<ITfSourceSingle>();

    THROW_IF_FAILED(spSrcTIM->AdviseSink(IID_ITfInputProcessorProfileActivationSink, static_cast<ITfInputProcessorProfileActivationSink*>(this), &_dwActivationSinkCookie));
    THROW_IF_FAILED(spSrcTIM->AdviseSink(IID_ITfUIElementSink, static_cast<ITfUIElementSink*>(this), &_dwUIElementSinkCookie));
    THROW_IF_FAILED(spSrcIC->AdviseSink(IID_ITfTextEditSink, static_cast<ITfTextEditSink*>(this), &_dwTextEditSinkCookie));
    THROW_IF_FAILED(spSrcICS->AdviseSingleSink(_tid, IID_ITfCleanupContextSink, static_cast<ITfCleanupContextSink*>(this)));
}

CConsoleTSF::~CConsoleTSF()
{
    // Destroy the current conversion area object

    if (_pConversionArea)
    {
        delete _pConversionArea;
        _pConversionArea = nullptr;
    }

    // Detach Cicero event sinks.
    if (_spITfInputContext)
    {
        const auto spSrcICS = _spITfInputContext.try_query<ITfSourceSingle>();
        if (spSrcICS)
        {
            spSrcICS->UnadviseSingleSink(_tid, IID_ITfCleanupContextSink);
        }
    }

    // Associate the document\context with the console window.

    if (_spITfThreadMgr)
    {
        const auto spSrcTIM = _spITfThreadMgr.try_query<ITfSource>();
        if (spSrcTIM)
        {
            if (_dwUIElementSinkCookie)
            {
                spSrcTIM->UnadviseSink(_dwUIElementSinkCookie);
            }
            if (_dwActivationSinkCookie)
            {
                spSrcTIM->UnadviseSink(_dwActivationSinkCookie);
            }
        }
    }

    _dwUIElementSinkCookie = 0;
    _dwActivationSinkCookie = 0;

    if (_spITfInputContext)
    {
        const auto spSrcIC = _spITfInputContext.try_query<ITfSource>();
        if (spSrcIC)
        {
            if (_dwContextOwnerCookie)
            {
                spSrcIC->UnadviseSink(_dwContextOwnerCookie);
            }
            if (_dwTextEditSinkCookie)
            {
                spSrcIC->UnadviseSink(_dwTextEditSinkCookie);
            }
        }
    }
    _dwContextOwnerCookie = 0;
    _dwTextEditSinkCookie = 0;

    // Clear the Cicero reference to our document manager.

    if (_spITfThreadMgr && _spITfDocumentMgr)
    {
        wil::com_ptr<ITfDocumentMgr> spDocMgr;
        _spITfThreadMgr->AssociateFocus(_hwndConsole, nullptr, &spDocMgr);
    }

    // Dismiss the input context and document manager.

    if (_spITfDocumentMgr)
    {
        _spITfDocumentMgr->Pop(TF_POPF_ALL);
    }

    _spITfInputContext.reset();
    _spITfDocumentMgr.reset();

    // Deactivate per-thread Cicero and uninitialize COM.

    if (_spITfThreadMgr)
    {
        _spITfThreadMgr->Deactivate();
        _spITfThreadMgr.reset();
    }
}

//+---------------------------------------------------------------------------
//
// CConsoleTSF::IUnknown::QueryInterface
// CConsoleTSF::IUnknown::AddRef
// CConsoleTSF::IUnknown::Release
//
//----------------------------------------------------------------------------

STDMETHODIMP CConsoleTSF::QueryInterface(REFIID riid, void** ppvObj)
{
    if (!ppvObj)
    {
        return E_POINTER;
    }

    if (IsEqualGUID(riid, IID_ITfCleanupContextSink))
    {
        *ppvObj = static_cast<ITfCleanupContextSink*>(this);
    }
    else if (IsEqualGUID(riid, IID_ITfContextOwnerCompositionSink))
    {
        *ppvObj = static_cast<ITfContextOwnerCompositionSink*>(this);
    }
    else if (IsEqualGUID(riid, IID_ITfUIElementSink))
    {
        *ppvObj = static_cast<ITfUIElementSink*>(this);
    }
    else if (IsEqualGUID(riid, IID_ITfContextOwner))
    {
        *ppvObj = static_cast<ITfContextOwner*>(this);
    }
    else if (IsEqualGUID(riid, IID_ITfInputProcessorProfileActivationSink))
    {
        *ppvObj = static_cast<ITfInputProcessorProfileActivationSink*>(this);
    }
    else if (IsEqualGUID(riid, IID_ITfTextEditSink))
    {
        *ppvObj = static_cast<ITfTextEditSink*>(this);
    }
    else if (IsEqualGUID(riid, IID_IUnknown))
    {
        *ppvObj = static_cast<IUnknown*>(static_cast<ITfContextOwner*>(this));
    }
    else
    {
        *ppvObj = nullptr;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

ULONG STDMETHODCALLTYPE CConsoleTSF::AddRef()
{
    return InterlockedIncrement(&_referenceCount);
}

ULONG STDMETHODCALLTYPE CConsoleTSF::Release()
{
    const auto cr = InterlockedDecrement(&_referenceCount);
    if (cr == 0)
    {
        delete this;
    }
    return cr;
}

//+---------------------------------------------------------------------------
//
// CConsoleTSF::ITfCleanupContextSink::OnCleanupContext
//
//----------------------------------------------------------------------------

STDMETHODIMP CConsoleTSF::OnCleanupContext(TfEditCookie ecWrite, ITfContext* pic)
{
    //
    // Remove GUID_PROP_COMPOSING
    //
    wil::com_ptr<ITfProperty> prop;
    if (SUCCEEDED(pic->GetProperty(GUID_PROP_COMPOSING, &prop)))
    {
        wil::com_ptr<IEnumTfRanges> enumranges;
        if (SUCCEEDED(prop->EnumRanges(ecWrite, &enumranges, nullptr)))
        {
            wil::com_ptr<ITfRange> rangeTmp;
            while (enumranges->Next(1, &rangeTmp, nullptr) == S_OK)
            {
                VARIANT var;
                VariantInit(&var);
                prop->GetValue(ecWrite, rangeTmp.get(), &var);
                if ((var.vt == VT_I4) && (var.lVal != 0))
                {
                    prop->Clear(ecWrite, rangeTmp.get());
                }
            }
        }
    }
    return S_OK;
}

//+---------------------------------------------------------------------------
//
// CConsoleTSF::ITfContextOwnerCompositionSink::OnStartComposition
// CConsoleTSF::ITfContextOwnerCompositionSink::OnUpdateComposition
// CConsoleTSF::ITfContextOwnerCompositionSink::OnEndComposition
//
//----------------------------------------------------------------------------

STDMETHODIMP CConsoleTSF::OnStartComposition(ITfCompositionView* pCompView, BOOL* pfOk)
{
    if (!_pConversionArea || (_cCompositions > 0 && (!_fModifyingDoc)))
    {
        *pfOk = FALSE;
    }
    else
    {
        *pfOk = TRUE;
        // Ignore compositions triggered by our own edit sessions
        // (i.e. when the application is the composition owner)
        auto clsidCompositionOwner = GUID_APPLICATION;
        pCompView->GetOwnerClsid(&clsidCompositionOwner);
        if (!IsEqualGUID(clsidCompositionOwner, GUID_APPLICATION))
        {
            _cCompositions++;
            if (_cCompositions == 1)
            {
                LOG_IF_FAILED(ImeStartComposition());
            }
        }
    }
    return S_OK;
}

STDMETHODIMP CConsoleTSF::OnUpdateComposition(ITfCompositionView* /*pComp*/, ITfRange*)
{
    return S_OK;
}

STDMETHODIMP CConsoleTSF::OnEndComposition(ITfCompositionView* pCompView)
{
    if (!_cCompositions || !_pConversionArea)
    {
        return E_FAIL;
    }
    // Ignore compositions triggered by our own edit sessions
    // (i.e. when the application is the composition owner)
    auto clsidCompositionOwner = GUID_APPLICATION;
    pCompView->GetOwnerClsid(&clsidCompositionOwner);
    if (!IsEqualGUID(clsidCompositionOwner, GUID_APPLICATION))
    {
        _cCompositions--;
        if (!_cCompositions)
        {
            LOG_IF_FAILED(_OnCompleteComposition());
            LOG_IF_FAILED(ImeEndComposition());
        }
    }
    return S_OK;
}

//+---------------------------------------------------------------------------
//
// CConsoleTSF::ITfTextEditSink::OnEndEdit
//
//----------------------------------------------------------------------------

STDMETHODIMP CConsoleTSF::OnEndEdit(ITfContext* pInputContext, TfEditCookie ecReadOnly, ITfEditRecord* pEditRecord)
{
    if (_cCompositions && _pConversionArea && _HasCompositionChanged(pInputContext, ecReadOnly, pEditRecord))
    {
        LOG_IF_FAILED(_OnUpdateComposition());
    }
    return S_OK;
}

//+---------------------------------------------------------------------------
//
// CConsoleTSF::ITfInputProcessorProfileActivationSink::OnActivated
//
//----------------------------------------------------------------------------

STDMETHODIMP CConsoleTSF::OnActivated(DWORD /*dwProfileType*/, LANGID /*langid*/, REFCLSID /*clsid*/, REFGUID catid, REFGUID /*guidProfile*/, HKL /*hkl*/, DWORD dwFlags)
{
    if (!(dwFlags & TF_IPSINK_FLAG_ACTIVE))
    {
        return S_OK;
    }
    if (!IsEqualGUID(catid, GUID_TFCAT_TIP_KEYBOARD))
    {
        // Don't care for non-keyboard profiles.
        return S_OK;
    }

    try
    {
        CreateConversionArea();
    }
    CATCH_RETURN();

    return S_OK;
}

//+---------------------------------------------------------------------------
//
// CConsoleTSF::ITfUIElementSink::BeginUIElement
//
//----------------------------------------------------------------------------

STDMETHODIMP CConsoleTSF::BeginUIElement(DWORD /*dwUIElementId*/, BOOL* pbShow)
{
    *pbShow = TRUE;
    return S_OK;
}

//+---------------------------------------------------------------------------
//
// CConsoleTSF::ITfUIElementSink::UpdateUIElement
//
//----------------------------------------------------------------------------

STDMETHODIMP CConsoleTSF::UpdateUIElement(DWORD /*dwUIElementId*/)
{
    return S_OK;
}

//+---------------------------------------------------------------------------
//
// CConsoleTSF::ITfUIElementSink::EndUIElement
//
//----------------------------------------------------------------------------

STDMETHODIMP CConsoleTSF::EndUIElement(DWORD /*dwUIElementId*/)
{
    return S_OK;
}

//+---------------------------------------------------------------------------
//
// CConsoleTSF::CreateConversionAreaService
//
//----------------------------------------------------------------------------

CConversionArea* CConsoleTSF::CreateConversionArea()
{
    const bool fHadConvArea = (_pConversionArea != nullptr);

    if (!_pConversionArea)
    {
        _pConversionArea = new CConversionArea();
    }

    // Associate the document\context with the console window.
    if (!fHadConvArea)
    {
        wil::com_ptr<ITfDocumentMgr> spPrevDocMgr;
        _spITfThreadMgr->AssociateFocus(_hwndConsole, _pConversionArea ? _spITfDocumentMgr.get() : nullptr, &spPrevDocMgr);
    }

    return _pConversionArea;
}

//+---------------------------------------------------------------------------
//
// CConsoleTSF::OnUpdateComposition()
//
//----------------------------------------------------------------------------

[[nodiscard]] HRESULT CConsoleTSF::_OnUpdateComposition()
{
    if (_fEditSessionRequested)
    {
        return S_FALSE;
    }

    auto hr = E_OUTOFMEMORY;
    const auto pEditSession = new (std::nothrow) CEditSessionUpdateCompositionString(this);
    if (pEditSession)
    {
        // Can't use TF_ES_SYNC because called from OnEndEdit.
        _fEditSessionRequested = TRUE;
        _spITfInputContext->RequestEditSession(_tid, pEditSession, TF_ES_READWRITE, &hr);
        if (FAILED(hr))
        {
            pEditSession->Release();
            _fEditSessionRequested = FALSE;
        }
    }
    return hr;
}

//+---------------------------------------------------------------------------
//
// CConsoleTSF::OnCompleteComposition()
//
//----------------------------------------------------------------------------

[[nodiscard]] HRESULT CConsoleTSF::_OnCompleteComposition()
{
    // Update the composition area.

    auto hr = E_OUTOFMEMORY;
    if (const auto pEditSession = new (std::nothrow) CEditSessionCompositionComplete(this))
    {
        // The composition could have been finalized because of a caret move, therefore it must be
        // inserted synchronously while at the original caret position.(TF_ES_SYNC is ok for a nested RO session).
        _spITfInputContext->RequestEditSession(_tid, pEditSession, TF_ES_READ | TF_ES_SYNC, &hr);
        if (FAILED(hr))
        {
            pEditSession->Release();
        }
    }

    // Cleanup (empty the context range) after the last composition, unless a new one has started.
    if (!_fCleanupSessionRequested)
    {
        _fCleanupSessionRequested = TRUE;
        if (const auto pEditSessionCleanup = new (std::nothrow) CEditSessionCompositionCleanup(this))
        {
            // Can't use TF_ES_SYNC because requesting RW while called within another session.
            // For the same reason, must use explicit TF_ES_ASYNC, or the request will be rejected otherwise.
            _spITfInputContext->RequestEditSession(_tid, pEditSessionCleanup, TF_ES_READWRITE | TF_ES_ASYNC, &hr);
            if (FAILED(hr))
            {
                pEditSessionCleanup->Release();
                _fCleanupSessionRequested = FALSE;
            }
        }
    }
    return hr;
}

bool CConsoleTSF::_HasCompositionChanged(ITfContext* pInputContext, TfEditCookie ecReadOnly, ITfEditRecord* pEditRecord)
{
    BOOL fChanged;
    if (SUCCEEDED(pEditRecord->GetSelectionStatus(&fChanged)))
    {
        if (fChanged)
        {
            return TRUE;
        }
    }

    //
    // Find GUID_PROP_CONIME_TRACKCOMPOSITION property.
    //

    wil::com_ptr<ITfProperty> Property;
    wil::com_ptr<ITfRange> FoundRange;
    wil::com_ptr<ITfProperty> PropertyTrackComposition;

    auto bFound = FALSE;

    if (SUCCEEDED(pInputContext->GetProperty(GUID_PROP_CONIME_TRACKCOMPOSITION, &Property)))
    {
        wil::com_ptr<IEnumTfRanges> EnumFindFirstTrackCompRange;

        if (SUCCEEDED(Property->EnumRanges(ecReadOnly, &EnumFindFirstTrackCompRange, NULL)))
        {
            HRESULT hr;
            wil::com_ptr<ITfRange> range;

            while ((hr = EnumFindFirstTrackCompRange->Next(1, &range, nullptr)) == S_OK)
            {
                VARIANT var;
                VariantInit(&var);

                hr = Property->GetValue(ecReadOnly, range.get(), &var);
                if (SUCCEEDED(hr))
                {
                    if ((V_VT(&var) == VT_I4 && V_I4(&var) != 0))
                    {
                        range->Clone(&FoundRange);
                        bFound = TRUE; // FOUND!!
                        break;
                    }
                }

                VariantClear(&var);

                if (bFound)
                {
                    break; // FOUND!!
                }
            }
        }
    }

    //
    // if there is no track composition property,
    // the composition has been changed since we put it.
    //
    if (!bFound)
    {
        return TRUE;
    }

    if (FoundRange == nullptr)
    {
        return FALSE;
    }

    bFound = FALSE; // RESET bFound flag...

    wil::com_ptr<ITfRange> rangeTrackComposition;
    if (SUCCEEDED(FoundRange->Clone(&rangeTrackComposition)))
    {
        //
        // get the text range that does not include read only area for
        // reconversion.
        //
        wil::com_ptr<ITfRange> rangeAllText;
        LONG cch;
        if (SUCCEEDED(CEditSessionObject::GetAllTextRange(ecReadOnly, pInputContext, &rangeAllText, &cch)))
        {
            LONG lResult;
            if (SUCCEEDED(rangeTrackComposition->CompareStart(ecReadOnly, rangeAllText.get(), TF_ANCHOR_START, &lResult)))
            {
                //
                // if the start position of the track composition range is not
                // the beginning of IC,
                // the composition has been changed since we put it.
                //
                if (lResult != 0)
                {
                    bFound = TRUE; // FOUND!!
                }
                else if (SUCCEEDED(rangeTrackComposition->CompareEnd(ecReadOnly, rangeAllText.get(), TF_ANCHOR_END, &lResult)))
                {
                    //
                    // if the start position of the track composition range is not
                    // the beginning of IC,
                    // the composition has been changed since we put it.
                    //
                    //
                    // If we find the changes in these property, we need to update hIMC.
                    //
                    const GUID* guids[] = { &GUID_PROP_COMPOSING,
                                            &GUID_PROP_ATTRIBUTE };
                    const int guid_size = sizeof(guids) / sizeof(GUID*);

                    wil::com_ptr<IEnumTfRanges> EnumPropertyChanged;

                    if (lResult != 0)
                    {
                        bFound = TRUE; // FOUND!!
                    }
                    else if (SUCCEEDED(pEditRecord->GetTextAndPropertyUpdates(TF_GTP_INCL_TEXT, guids, guid_size, &EnumPropertyChanged)))
                    {
                        HRESULT hr;
                        wil::com_ptr<ITfRange> range;

                        while ((hr = EnumPropertyChanged->Next(1, &range, nullptr)) == S_OK)
                        {
                            BOOL empty;
                            if (range->IsEmpty(ecReadOnly, &empty) == S_OK && empty)
                            {
                                continue;
                            }

                            bFound = TRUE; // FOUND!!
                            break;
                        }
                    }
                }
            }
        }
    }
    return bFound;
}
