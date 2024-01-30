// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

#include "precomp.h"

static CConsoleTSF* g_pConsoleTSF = nullptr;

extern "C" BOOL ActivateTextServices(HWND hwndConsole, GetSuggestionWindowPos pfnPosition, GetTextBoxAreaPos pfnTextArea)
{
    if (!g_pConsoleTSF && hwndConsole)
    {
        g_pConsoleTSF = new CConsoleTSF(hwndConsole, pfnPosition, pfnTextArea);
        if (g_pConsoleTSF)
        {
            // Conhost calls this function only when the console window has focus.
            g_pConsoleTSF->SetFocus(TRUE);
        }
        else
        {
            SafeReleaseClear(g_pConsoleTSF);
        }
    }
    return g_pConsoleTSF ? TRUE : FALSE;
}

extern "C" void DeactivateTextServices()
{
    if (g_pConsoleTSF)
    {
        g_pConsoleTSF->Release();
        g_pConsoleTSF = nullptr;
    }
}
