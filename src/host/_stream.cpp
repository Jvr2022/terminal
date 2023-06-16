// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

#include "precomp.h"

#include "ApiRoutines.h"

#include "_stream.h"
#include "stream.h"
#include "writeData.hpp"

#include "_output.h"
#include "output.h"
#include "dbcs.h"
#include "handle.h"
#include "misc.h"

#include "../types/inc/convert.hpp"
#include "../types/inc/GlyphWidth.hpp"
#include "../types/inc/Viewport.hpp"

#include "../interactivity/inc/ServiceLocator.hpp"

#pragma hdrstop
using namespace Microsoft::Console::Types;
using Microsoft::Console::Interactivity::ServiceLocator;
using Microsoft::Console::VirtualTerminal::StateMachine;
// Used by WriteCharsLegacy.
#define IS_GLYPH_CHAR(wch) (((wch) >= L' ') && ((wch) != 0x007F))

// Routine Description:
// - This routine updates the cursor position.  Its input is the non-special
//   cased new location of the cursor.  For example, if the cursor were being
//   moved one space backwards from the left edge of the screen, the X
//   coordinate would be -1.  This routine would set the X coordinate to
//   the right edge of the screen and decrement the Y coordinate by one.
// Arguments:
// - screenInfo - reference to screen buffer information structure.
// - coordCursor - New location of cursor.
// - fKeepCursorVisible - TRUE if changing window origin desirable when hit right edge
// Return Value:
void AdjustCursorPosition(SCREEN_INFORMATION& screenInfo, _In_ til::point coordCursor, const BOOL fKeepCursorVisible, _Inout_opt_ til::CoordType* psScrollY)
{
    const auto bufferSize = screenInfo.GetBufferSize().Dimensions();
    if (coordCursor.x < 0)
    {
        if (coordCursor.y > 0)
        {
            coordCursor.x = bufferSize.width + coordCursor.x;
            coordCursor.y = coordCursor.y - 1;
        }
        else
        {
            coordCursor.x = 0;
        }
    }
    else if (coordCursor.x >= bufferSize.width)
    {
        // at end of line. if wrap mode, wrap cursor.  otherwise leave it where it is.
        if (screenInfo.OutputMode & ENABLE_WRAP_AT_EOL_OUTPUT)
        {
            coordCursor.y += coordCursor.x / bufferSize.width;
            coordCursor.x = coordCursor.x % bufferSize.width;
        }
        else
        {
            coordCursor.x = screenInfo.GetTextBuffer().GetCursor().GetPosition().x;
        }
    }

    if (coordCursor.y >= bufferSize.height)
    {
        // At the end of the buffer. Scroll contents of screen buffer so new position is visible.
        StreamScrollRegion(screenInfo);

        if (nullptr != psScrollY)
        {
            *psScrollY += bufferSize.height - coordCursor.y - 1;
        }
        coordCursor.y += bufferSize.height - coordCursor.y - 1;
    }

    const auto cursorMovedPastViewport = coordCursor.y > screenInfo.GetViewport().BottomInclusive();

    // if at right or bottom edge of window, scroll right or down one char.
    if (cursorMovedPastViewport)
    {
        til::point WindowOrigin;
        WindowOrigin.x = 0;
        WindowOrigin.y = coordCursor.y - screenInfo.GetViewport().BottomInclusive();
        LOG_IF_FAILED(screenInfo.SetViewportOrigin(false, WindowOrigin, true));
    }

    if (fKeepCursorVisible)
    {
        screenInfo.MakeCursorVisible(coordCursor);
    }
    LOG_IF_FAILED(screenInfo.SetCursorPosition(coordCursor, !!fKeepCursorVisible));
}

// As the name implies, this writes text without processing its control characters.
static size_t _writeCharsLegacyUnprocessed(SCREEN_INFORMATION& screenInfo, const DWORD dwFlags, til::CoordType* const psScrollY, const std::wstring_view& text)
{
    const auto fKeepCursorVisible = WI_IsFlagSet(dwFlags, WC_KEEP_CURSOR_VISIBLE);
    const auto fWrapAtEOL = WI_IsFlagSet(screenInfo.OutputMode, ENABLE_WRAP_AT_EOL_OUTPUT);
    const auto hasAccessibilityEventing = screenInfo.HasAccessibilityEventing();
    auto& textBuffer = screenInfo.GetTextBuffer();
    auto CursorPosition = textBuffer.GetCursor().GetPosition();
    size_t TempNumSpaces = 0;

    RowWriteState state{
        .text = text,
        .columnLimit = til::CoordTypeMax,
    };

    while (!state.text.empty())
    {
        state.columnBegin = CursorPosition.x;
        textBuffer.WriteLine(CursorPosition.y, fWrapAtEOL, textBuffer.GetCurrentAttributes(), state);
        CursorPosition.x = state.columnEnd;

        const auto spaces = gsl::narrow_cast<size_t>(state.columnEnd - state.columnBegin);
        TempNumSpaces += spaces;

        if (hasAccessibilityEventing && spaces != 0)
        {
            screenInfo.NotifyAccessibilityEventing(state.columnBegin, CursorPosition.y, state.columnEnd - 1, CursorPosition.y);
        }

        AdjustCursorPosition(screenInfo, CursorPosition, fKeepCursorVisible, psScrollY);
        CursorPosition = textBuffer.GetCursor().GetPosition();
    }

    return TempNumSpaces;
}

// Routine Description:
// - This routine writes a string to the screen, processing any embedded
//   unicode characters.  The string is also copied to the input buffer, if
//   the output mode is line mode.
// Arguments:
// - screenInfo - reference to screen buffer information structure.
// - pwchBufferBackupLimit - Pointer to beginning of buffer.
// - pwchBuffer - Pointer to buffer to copy string to.  assumed to be at least as long as pwchRealUnicode.
//                This pointer is updated to point to the next position in the buffer.
// - pwchRealUnicode - Pointer to string to write.
// - pcb - On input, number of bytes to write.  On output, number of bytes written.
// - pcSpaces - On output, the number of spaces consumed by the written characters.
// - dwFlags -
//      WC_INTERACTIVE             backspace overwrites characters, control characters are expanded (as in, to "^X")
//      WC_KEEP_CURSOR_VISIBLE     change window origin desirable when hit rt. edge
// Return Value:
// Note:
// - This routine does not process tabs and backspace properly.  That code will be implemented as part of the line editing services.
[[nodiscard]] NTSTATUS WriteCharsLegacy(SCREEN_INFORMATION& screenInfo,
                                        _In_range_(<=, pwchBuffer) const wchar_t* const pwchBufferBackupLimit,
                                        _In_ const wchar_t* pwchBuffer,
                                        _In_reads_bytes_(*pcb) const wchar_t* pwchRealUnicode,
                                        _Inout_ size_t* const pcb,
                                        _Out_opt_ size_t* const pcSpaces,
                                        const til::CoordType sOriginalXPosition,
                                        const DWORD dwFlags,
                                        _Inout_opt_ til::CoordType* const psScrollY)
try
{
    static constexpr wchar_t tabSpaces[8]{ L' ', L' ', L' ', L' ', L' ', L' ', L' ', L' ' };

    UNREFERENCED_PARAMETER(sOriginalXPosition);
    UNREFERENCED_PARAMETER(pwchBuffer);
    UNREFERENCED_PARAMETER(pwchBufferBackupLimit);

    auto& textBuffer = screenInfo.GetTextBuffer();
    auto& cursor = textBuffer.GetCursor();
    const auto width = textBuffer.GetSize().Width();
    const auto Attributes = textBuffer.GetCurrentAttributes();
    const auto fKeepCursorVisible = WI_IsFlagSet(dwFlags, WC_KEEP_CURSOR_VISIBLE);
    const auto fUnprocessed = WI_IsFlagClear(screenInfo.OutputMode, ENABLE_PROCESSED_OUTPUT);
    const auto fWrapAtEOL = WI_IsFlagSet(screenInfo.OutputMode, ENABLE_WRAP_AT_EOL_OUTPUT);
    auto it = pwchRealUnicode;
    const auto end = it + *pcb / sizeof(wchar_t);
    size_t TempNumSpaces = 0;

    if (cursor.IsDelayedEOLWrap() && fWrapAtEOL)
    {
        auto CursorPosition = cursor.GetPosition();
        const auto delayedCursorPosition = cursor.GetDelayedAtPosition();
        cursor.ResetDelayEOLWrap();
        if (delayedCursorPosition == CursorPosition)
        {
            CursorPosition.x = 0;
            CursorPosition.y++;
            AdjustCursorPosition(screenInfo, CursorPosition, fKeepCursorVisible, psScrollY);
        }
    }

    if (fUnprocessed)
    {
        TempNumSpaces += _writeCharsLegacyUnprocessed(screenInfo, dwFlags, psScrollY, { it, end });
        it = end;
    }

    while (it != end)
    {
        const auto nextControlChar = std::find_if(it, end, [](const auto& wch) { return !IS_GLYPH_CHAR(wch); });
        if (nextControlChar != it)
        {
            TempNumSpaces += _writeCharsLegacyUnprocessed(screenInfo, dwFlags, psScrollY, { it, nextControlChar });
            it = nextControlChar;
        }

        for (; it != end && !IS_GLYPH_CHAR(*it); ++it)
        {
            switch (*it)
            {
            case UNICODE_NULL:
                if (WI_IsFlagSet(dwFlags, WC_INTERACTIVE))
                {
                    break;
                }
                TempNumSpaces += _writeCharsLegacyUnprocessed(screenInfo, dwFlags, psScrollY, { &tabSpaces[0], 1 });
                continue;
            case UNICODE_BELL:
                if (WI_IsFlagSet(dwFlags, WC_INTERACTIVE))
                {
                    break;
                }
                std::ignore = screenInfo.SendNotifyBeep();
                continue;
            case UNICODE_BACKSPACE:
            {
                auto CursorPosition = cursor.GetPosition();

                if (WI_IsFlagClear(dwFlags, WC_INTERACTIVE))
                {
                    CursorPosition.x = textBuffer.GetRowByOffset(CursorPosition.y).NavigateToPrevious(CursorPosition.x);
                    AdjustCursorPosition(screenInfo, CursorPosition, fKeepCursorVisible, psScrollY);
                    continue;
                }

                const auto moveUp = [&]() {
                    CursorPosition.x = -1;
                    AdjustCursorPosition(screenInfo, CursorPosition, fKeepCursorVisible, psScrollY);

                    const auto y = cursor.GetPosition().y;
                    auto& row = textBuffer.GetRowByOffset(y);

                    CursorPosition.x = width;
                    CursorPosition.y = y;

                    if (row.WasDoubleBytePadded())
                    {
                        CursorPosition.x--;
                        TempNumSpaces--;
                    }

                    row.SetWrapForced(false);
                    row.SetDoubleBytePadded(false);
                };

                // We have to move up early because the tab handling code below needs
                // to be on the row of the tab already, so that we can call GetText().
                if (CursorPosition.x == 0 && CursorPosition.y != 0)
                {
                    moveUp();
                }

                til::CoordType glyphCount = 1;

                if (pwchBuffer != pwchBufferBackupLimit)
                {
                    const auto LastChar = pwchBuffer[-1];

                    // Deleting tabs is a bit tricky, because they have a variable width between 1 and 8 spaces,
                    // are stored as whitespace but are technically distinct from whitespace.
                    if (LastChar == UNICODE_TAB)
                    {
                        const auto precedingText = textBuffer.GetRowByOffset(CursorPosition.y).GetText(CursorPosition.x - 8, CursorPosition.x);

                        // First, we measure the amount of spaces that precede the cursor in the text buffer,
                        // which is generally the amount of spaces that we end up deleting. We do it this way,
                        // because we don't know what kind of complex mix of wide/narrow glyphs precede the tab.
                        // Basically, by asking the text buffer we get the size information of the preceding text.
                        if (precedingText.size() >= 2 && precedingText.back() == L' ')
                        {
                            auto textIt = precedingText.rbegin() + 1;
                            const auto textEnd = precedingText.rend();

                            for (; textIt != textEnd && *textIt == L' '; ++textIt)
                            {
                                glyphCount++;
                            }
                        }

                        // But there's a problem: When you print "  \t" it should delete 6 spaces and not 8.
                        // In other words, we shouldn't delete any actual preceding whitespaces. We can ask
                        // the "backup" buffer (= preceding text in the commandline) for this information.
                        const auto backupEnd = pwchBuffer - 1;
                        const auto backupLimit = backupEnd - std::min<ptrdiff_t>(7, backupEnd - pwchBufferBackupLimit);
                        auto backupBeg = backupEnd;
                        for (; backupBeg != backupLimit && backupBeg[-1] == L' '; --backupBeg)
                        {
                        }
                        glyphCount -= std::min(glyphCount - 1, gsl::narrow_cast<til::CoordType>(backupEnd - backupBeg));
                    }
                    // Control chars in interactive mode where previously written out
                    // as ^X for instance, so now we also need to delete 2 glyphs.
                    else if (IS_CONTROL_CHAR(LastChar))
                    {
                        glyphCount = 2;
                    }
                }

                for (;;)
                {
                    // We've already moved up if the cursor was in the first column so
                    // we need to start off with overwriting the text with whitespace.
                    // It wouldn't make sense to check the cursor position again already.
                    {
                        const auto previousColumn = CursorPosition.x;
                        CursorPosition.x = textBuffer.GetRowByOffset(CursorPosition.y).NavigateToPrevious(previousColumn);

                        RowWriteState state{
                            .text = { &tabSpaces[0], 8 },
                            .columnBegin = CursorPosition.x,
                            .columnLimit = previousColumn,
                        };
                        textBuffer.WriteLine(CursorPosition.y, fWrapAtEOL, Attributes, state);
                        TempNumSpaces -= previousColumn - CursorPosition.x;
                    }

                    // The cursor movement logic is a little different for the last iteration, so we exit early here.
                    glyphCount--;
                    if (glyphCount <= 0)
                    {
                        break;
                    }

                    // Otherwise, in case we need to delete 2 or more glyphs, we need to check it again.
                    if (CursorPosition.x == 0 && CursorPosition.y != 0)
                    {
                        moveUp();
                    }
                }

                // After the last iteration the cursor might now be in the first column after a line
                // that was previously padded with a whitespace in the last column due to a wide glyph.
                // Now that the wide glyph is presumably gone, we can move up a line.
                if (CursorPosition.x == 0 && CursorPosition.y != 0 && textBuffer.GetRowByOffset(CursorPosition.y - 1).WasDoubleBytePadded())
                {
                    moveUp();
                }
                else
                {
                    AdjustCursorPosition(screenInfo, CursorPosition, fKeepCursorVisible, psScrollY);
                }

                // Notify accessibility to read the backspaced character.
                // See GH:12735, MSFT:31748387
                if (screenInfo.HasAccessibilityEventing())
                {
                    if (auto pConsoleWindow = ServiceLocator::LocateConsoleWindow())
                    {
                        LOG_IF_FAILED(pConsoleWindow->SignalUia(UIA_Text_TextChangedEventId));
                    }
                }
                continue;
            }
            case UNICODE_TAB:
            {
                auto CursorPosition = cursor.GetPosition();
                const auto tabCount = gsl::narrow_cast<size_t>(NUMBER_OF_SPACES_IN_TAB(CursorPosition.x));
                TempNumSpaces += _writeCharsLegacyUnprocessed(screenInfo, dwFlags, psScrollY, { &tabSpaces[0], tabCount });
                continue;
            }
            case UNICODE_LINEFEED:
            {
                auto CursorPosition = cursor.GetPosition();
                if (WI_IsFlagClear(screenInfo.OutputMode, DISABLE_NEWLINE_AUTO_RETURN))
                {
                    CursorPosition.x = 0;
                }

                textBuffer.GetRowByOffset(CursorPosition.y).SetWrapForced(false);
                CursorPosition.y = CursorPosition.y + 1;
                AdjustCursorPosition(screenInfo, CursorPosition, fKeepCursorVisible, psScrollY);
                continue;
            }
            case UNICODE_CARRIAGERETURN:
            {
                auto CursorPosition = cursor.GetPosition();
                CursorPosition.x = 0;
                AdjustCursorPosition(screenInfo, CursorPosition, fKeepCursorVisible, psScrollY);
                continue;
            }
            default:
                break;
            }

            if (WI_IsFlagSet(dwFlags, WC_INTERACTIVE) && IS_CONTROL_CHAR(*it))
            {
                const wchar_t wchs[2]{ L'^', static_cast<wchar_t>(*it + L'@') };
                TempNumSpaces += _writeCharsLegacyUnprocessed(screenInfo, dwFlags, psScrollY, { &wchs[0], 2 });
            }
            else
            {
                // TODO test
                // As a special favor to incompetent apps that attempt to display control chars,
                // convert to corresponding OEM Glyph Chars
                const auto cp = ServiceLocator::LocateGlobals().getConsoleInformation().OutputCP;
                wchar_t buffer;
                const auto result = MultiByteToWideChar(cp, MB_USEGLYPHCHARS, reinterpret_cast<LPCSTR>(it), 1, &buffer, 1);
                if (result == 1)
                {
                    TempNumSpaces += _writeCharsLegacyUnprocessed(screenInfo, dwFlags, psScrollY, { &buffer, 1 });
                }
            }
        }
    }

    if (pcSpaces)
    {
        *pcSpaces = TempNumSpaces;
    }

    return S_OK;
}
NT_CATCH_RETURN()

// Routine Description:
// - This routine writes a string to the screen, processing any embedded
//   unicode characters.  The string is also copied to the input buffer, if
//   the output mode is line mode.
// Arguments:
// - screenInfo - reference to screen buffer information structure.
// - pwchBufferBackupLimit - Pointer to beginning of buffer.
// - pwchBuffer - Pointer to buffer to copy string to.  assumed to be at least as long as pwchRealUnicode.
//              This pointer is updated to point to the next position in the buffer.
// - pwchRealUnicode - Pointer to string to write.
// - pcb - On input, number of bytes to write.  On output, number of bytes written.
// - pcSpaces - On output, the number of spaces consumed by the written characters.
// - dwFlags -
//      WC_INTERACTIVE             backspace overwrites characters, control characters are expanded (as in, to "^X")
//      WC_KEEP_CURSOR_VISIBLE     change window origin (viewport) desirable when hit rt. edge
// Return Value:
// Note:
// - This routine does not process tabs and backspace properly.  That code will be implemented as part of the line editing services.
[[nodiscard]] NTSTATUS WriteChars(SCREEN_INFORMATION& screenInfo,
                                  _In_range_(<=, pwchBuffer) const wchar_t* const pwchBufferBackupLimit,
                                  _In_ const wchar_t* pwchBuffer,
                                  _In_reads_bytes_(*pcb) const wchar_t* pwchRealUnicode,
                                  _Inout_ size_t* const pcb,
                                  _Out_opt_ size_t* const pcSpaces,
                                  const til::CoordType sOriginalXPosition,
                                  const DWORD dwFlags,
                                  _Inout_opt_ til::CoordType* const psScrollY)
try
{
    if (WI_IsAnyFlagClear(screenInfo.OutputMode, ENABLE_VIRTUAL_TERMINAL_PROCESSING | ENABLE_PROCESSED_OUTPUT))
    {
        return WriteCharsLegacy(screenInfo,
                                pwchBufferBackupLimit,
                                pwchBuffer,
                                pwchRealUnicode,
                                pcb,
                                pcSpaces,
                                sOriginalXPosition,
                                dwFlags,
                                psScrollY);
    }

    auto& machine = screenInfo.GetStateMachine();
    const auto cch = *pcb / sizeof(WCHAR);

    machine.ProcessString({ pwchRealUnicode, cch });

    if (nullptr != pcSpaces)
    {
        *pcSpaces = 0;
    }

    return STATUS_SUCCESS;
}
NT_CATCH_RETURN()

// Routine Description:
// - Takes the given text and inserts it into the given screen buffer.
// Note:
// - Console lock must be held when calling this routine
// - String has been translated to unicode at this point.
// Arguments:
// - pwchBuffer - wide character text to be inserted into buffer
// - pcbBuffer - byte count of pwchBuffer on the way in, number of bytes consumed on the way out.
// - screenInfo - Screen Information class to write the text into at the current cursor position
// - ppWaiter - If writing to the console is blocked for whatever reason, this will be filled with a pointer to context
//              that can be used by the server to resume the call at a later time.
// Return Value:
// - STATUS_SUCCESS if OK.
// - CONSOLE_STATUS_WAIT if we couldn't finish now and need to be called back later (see ppWaiter).
// - Or a suitable NTSTATUS format error code for memory/string/math failures.
[[nodiscard]] NTSTATUS DoWriteConsole(_In_reads_bytes_(*pcbBuffer) PCWCHAR pwchBuffer,
                                      _Inout_ size_t* const pcbBuffer,
                                      SCREEN_INFORMATION& screenInfo,
                                      bool requiresVtQuirk,
                                      std::unique_ptr<WriteData>& waiter)
{
    const auto& gci = ServiceLocator::LocateGlobals().getConsoleInformation();
    if (WI_IsAnyFlagSet(gci.Flags, (CONSOLE_SUSPENDED | CONSOLE_SELECTING | CONSOLE_SCROLLBAR_TRACKING)))
    {
        try
        {
            waiter = std::make_unique<WriteData>(screenInfo,
                                                 pwchBuffer,
                                                 *pcbBuffer,
                                                 gci.OutputCP,
                                                 requiresVtQuirk);
        }
        catch (...)
        {
            return NTSTATUS_FROM_HRESULT(wil::ResultFromCaughtException());
        }

        return CONSOLE_STATUS_WAIT;
    }

    auto restoreVtQuirk{
        wil::scope_exit([&]() { screenInfo.ResetIgnoreLegacyEquivalentVTAttributes(); })
    };

    if (requiresVtQuirk)
    {
        screenInfo.SetIgnoreLegacyEquivalentVTAttributes();
    }
    else
    {
        restoreVtQuirk.release();
    }

    const auto& textBuffer = screenInfo.GetTextBuffer();
    return WriteChars(screenInfo,
                      pwchBuffer,
                      pwchBuffer,
                      pwchBuffer,
                      pcbBuffer,
                      nullptr,
                      textBuffer.GetCursor().GetPosition().x,
                      0,
                      nullptr);
}

// Routine Description:
// - This method performs the actual work of attempting to write to the console, converting data types as necessary
//   to adapt from the server types to the legacy internal host types.
// - It operates on Unicode data only. It's assumed the text is translated by this point.
// Arguments:
// - OutContext - the console output object to write the new text into
// - pwsTextBuffer - wide character text buffer provided by client application to insert
// - cchTextBufferLength - text buffer counted in characters
// - pcchTextBufferRead - character count of the number of characters we were able to insert before returning
// - ppWaiter - If we are blocked from writing now and need to wait, this is filled with contextual data for the server to restore the call later
// Return Value:
// - S_OK if successful.
// - S_OK if we need to wait (check if ppWaiter is not nullptr).
// - Or a suitable HRESULT code for math/string/memory failures.
[[nodiscard]] HRESULT WriteConsoleWImplHelper(IConsoleOutputObject& context,
                                              const std::wstring_view buffer,
                                              size_t& read,
                                              bool requiresVtQuirk,
                                              std::unique_ptr<WriteData>& waiter) noexcept
{
    try
    {
        // Set out variables in case we exit early.
        read = 0;
        waiter.reset();

        // Convert characters to bytes to give to DoWriteConsole.
        size_t cbTextBufferLength;
        RETURN_IF_FAILED(SizeTMult(buffer.size(), sizeof(wchar_t), &cbTextBufferLength));

        auto Status = DoWriteConsole(const_cast<wchar_t*>(buffer.data()), &cbTextBufferLength, context, requiresVtQuirk, waiter);

        // Convert back from bytes to characters for the resulting string length written.
        read = cbTextBufferLength / sizeof(wchar_t);

        if (Status == CONSOLE_STATUS_WAIT)
        {
            FAIL_FAST_IF_NULL(waiter.get());
            Status = STATUS_SUCCESS;
        }

        RETURN_NTSTATUS(Status);
    }
    CATCH_RETURN();
}

// Routine Description:
// - Writes non-Unicode formatted data into the given console output object.
// - This method will convert from the given input into wide characters before chain calling the wide character version of the function.
//   It uses the current Output Codepage for conversions (set via SetConsoleOutputCP).
// - NOTE: This may be blocked for various console states and will return a wait context pointer if necessary.
// Arguments:
// - context - the console output object to write the new text into
// - buffer - char/byte text buffer provided by client application to insert
// - read - character count of the number of characters (also bytes because A version) we were able to insert before returning
// - waiter - If we are blocked from writing now and need to wait, this is filled with contextual data for the server to restore the call later
// Return Value:
// - S_OK if successful.
// - S_OK if we need to wait (check if ppWaiter is not nullptr).
// - Or a suitable HRESULT code for math/string/memory failures.
[[nodiscard]] HRESULT ApiRoutines::WriteConsoleAImpl(IConsoleOutputObject& context,
                                                     const std::string_view buffer,
                                                     size_t& read,
                                                     bool requiresVtQuirk,
                                                     std::unique_ptr<IWaitRoutine>& waiter) noexcept
{
    try
    {
        // Ensure output variables are initialized.
        read = 0;
        waiter.reset();

        if (buffer.empty())
        {
            return S_OK;
        }

        LockConsole();
        auto unlock{ wil::scope_exit([&] { UnlockConsole(); }) };

        auto& screenInfo{ context.GetActiveBuffer() };
        const auto& consoleInfo{ ServiceLocator::LocateGlobals().getConsoleInformation() };
        const auto codepage{ consoleInfo.OutputCP };
        auto leadByteCaptured{ false };
        auto leadByteConsumed{ false };
        std::wstring wstr{};
        static til::u8state u8State{};

        // Convert our input parameters to Unicode
        if (codepage == CP_UTF8)
        {
            RETURN_IF_FAILED(til::u8u16(buffer, wstr, u8State));
            read = buffer.size();
        }
        else
        {
            // In case the codepage changes from UTF-8 to another,
            // we discard partials that might still be cached.
            u8State.reset();

            int mbPtrLength{};
            RETURN_IF_FAILED(SizeTToInt(buffer.size(), &mbPtrLength));

            // (buffer.size() + 2) I think because we might be shoving another unicode char
            // from screenInfo->WriteConsoleDbcsLeadByte in front
            // because we previously checked that buffer.size() fits into an int, +2 won't cause an overflow of size_t
            wstr.resize(buffer.size() + 2);

            auto wcPtr{ wstr.data() };
            auto mbPtr{ buffer.data() };
            size_t dbcsLength{};
            if (screenInfo.WriteConsoleDbcsLeadByte[0] != 0 && gsl::narrow_cast<byte>(*mbPtr) >= byte{ ' ' })
            {
                // there was a portion of a dbcs character stored from a previous
                // call so we take the 2nd half from mbPtr[0], put them together
                // and write the wide char to wcPtr[0]
                screenInfo.WriteConsoleDbcsLeadByte[1] = gsl::narrow_cast<byte>(*mbPtr);

                try
                {
                    const auto wFromComplemented{
                        ConvertToW(codepage, { reinterpret_cast<const char*>(screenInfo.WriteConsoleDbcsLeadByte), ARRAYSIZE(screenInfo.WriteConsoleDbcsLeadByte) })
                    };

                    FAIL_FAST_IF(wFromComplemented.size() != 1);
                    dbcsLength = sizeof(wchar_t);
                    wcPtr[0] = wFromComplemented.at(0);
                    mbPtr++;
                }
                catch (...)
                {
                    dbcsLength = 0;
                }

                // this looks weird to be always incrementing even if the conversion failed, but this is the
                // original behavior so it's left unchanged.
                wcPtr++;
                mbPtrLength--;

                // Note that we used a stored lead byte from a previous call in order to complete this write
                // Use this to offset the "number of bytes consumed" calculation at the end by -1 to account
                // for using a byte we had internally, not off the stream.
                leadByteConsumed = true;
            }

            screenInfo.WriteConsoleDbcsLeadByte[0] = 0;

            // if the last byte in mbPtr is a lead byte for the current code page,
            // save it for the next time this function is called and we can piece it
            // back together then
            if (mbPtrLength != 0 && CheckBisectStringA(const_cast<char*>(mbPtr), mbPtrLength, &consoleInfo.OutputCPInfo))
            {
                screenInfo.WriteConsoleDbcsLeadByte[0] = gsl::narrow_cast<byte>(mbPtr[mbPtrLength - 1]);
                mbPtrLength--;

                // Note that we captured a lead byte during this call, but won't actually draw it until later.
                // Use this to offset the "number of bytes consumed" calculation at the end by +1 to account
                // for taking a byte off the stream.
                leadByteCaptured = true;
            }

            if (mbPtrLength != 0)
            {
                // convert the remaining bytes in mbPtr to wide chars
                mbPtrLength = sizeof(wchar_t) * MultiByteToWideChar(codepage, 0, mbPtr, mbPtrLength, wcPtr, mbPtrLength);
            }

            wstr.resize((dbcsLength + mbPtrLength) / sizeof(wchar_t));
        }

        // Hold the specific version of the waiter locally so we can tinker with it if we have to store additional context.
        std::unique_ptr<WriteData> writeDataWaiter{};

        // Make the W version of the call
        size_t wcBufferWritten{};
        const auto hr{ WriteConsoleWImplHelper(screenInfo, wstr, wcBufferWritten, requiresVtQuirk, writeDataWaiter) };

        // If there is no waiter, process the byte count now.
        if (nullptr == writeDataWaiter.get())
        {
            // Calculate how many bytes of the original A buffer were consumed in the W version of the call to satisfy mbBufferRead.
            // For UTF-8 conversions, we've already returned this information above.
            if (CP_UTF8 != codepage)
            {
                size_t mbBufferRead{};

                // Start by counting the number of A bytes we used in printing our W string to the screen.
                try
                {
                    mbBufferRead = GetALengthFromW(codepage, { wstr.data(), wcBufferWritten });
                }
                CATCH_LOG();

                // If we captured a byte off the string this time around up above, it means we didn't feed
                // it into the WriteConsoleW above, and therefore its consumption isn't accounted for
                // in the count we just made. Add +1 to compensate.
                if (leadByteCaptured)
                {
                    mbBufferRead++;
                }

                // If we consumed an internally-stored lead byte this time around up above, it means that we
                // fed a byte into WriteConsoleW that wasn't a part of this particular call's request.
                // We need to -1 to compensate and tell the caller the right number of bytes consumed this request.
                if (leadByteConsumed)
                {
                    mbBufferRead--;
                }

                read = mbBufferRead;
            }
        }
        else
        {
            // If there is a waiter, then we need to stow some additional information in the wait structure so
            // we can synthesize the correct byte count later when the wait routine is triggered.
            if (CP_UTF8 != codepage)
            {
                // For non-UTF8 codepages, save the lead byte captured/consumed data so we can +1 or -1 the final decoded count
                // in the WaitData::Notify method later.
                writeDataWaiter->SetLeadByteAdjustmentStatus(leadByteCaptured, leadByteConsumed);
            }
            else
            {
                // For UTF8 codepages, just remember the consumption count from the UTF-8 parser.
                writeDataWaiter->SetUtf8ConsumedCharacters(read);
            }
        }

        // Give back the waiter now that we're done with tinkering with it.
        waiter.reset(writeDataWaiter.release());

        return hr;
    }
    CATCH_RETURN();
}

// Routine Description:
// - Writes Unicode formatted data into the given console output object.
// - NOTE: This may be blocked for various console states and will return a wait context pointer if necessary.
// Arguments:
// - OutContext - the console output object to write the new text into
// - pwsTextBuffer - wide character text buffer provided by client application to insert
// - cchTextBufferLength - text buffer counted in characters
// - pcchTextBufferRead - character count of the number of characters we were able to insert before returning
// - ppWaiter - If we are blocked from writing now and need to wait, this is filled with contextual data for the server to restore the call later
// Return Value:
// - S_OK if successful.
// - S_OK if we need to wait (check if ppWaiter is not nullptr).
// - Or a suitable HRESULT code for math/string/memory failures.
[[nodiscard]] HRESULT ApiRoutines::WriteConsoleWImpl(IConsoleOutputObject& context,
                                                     const std::wstring_view buffer,
                                                     size_t& read,
                                                     bool requiresVtQuirk,
                                                     std::unique_ptr<IWaitRoutine>& waiter) noexcept
{
    try
    {
        LockConsole();
        auto unlock = wil::scope_exit([&] { UnlockConsole(); });

        std::unique_ptr<WriteData> writeDataWaiter;
        RETURN_IF_FAILED(WriteConsoleWImplHelper(context.GetActiveBuffer(), buffer, read, requiresVtQuirk, writeDataWaiter));

        // Transfer specific waiter pointer into the generic interface wrapper.
        waiter.reset(writeDataWaiter.release());

        return S_OK;
    }
    CATCH_RETURN();
}
