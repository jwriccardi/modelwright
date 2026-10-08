using System;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>
/// Thrown by an <see cref="IPrecedentProvider"/> when Excel cannot answer right now (it is busy: a cell is being
/// edited, a dialog is open): nothing was read, and the same call can succeed later. <see cref="PrecedentTree"/> then
/// leaves the node unloaded, so expanding it again retries, instead of keeping an error row for it.
/// </summary>
public sealed class PrecedentsUnavailableException : Exception
{
    // Excel rejects a call while it is busy (VBA_E_IGNORE; also returned while a cell is being edited), or COM rejects
    // it (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER).
    private const int VbaIgnore = unchecked((int)0x800AC472);
    private const int CallRejected = unchecked((int)0x80010001);
    private const int RetryLater = unchecked((int)0x8001010A);

    /// <summary>Creates the exception with a default message.</summary>
    public PrecedentsUnavailableException()
        : this("Excel is busy (a cell is being edited or a dialog is open): try again.")
    {
    }

    /// <summary>Creates the exception.</summary>
    public PrecedentsUnavailableException(string message)
        : base(message)
    {
    }

    /// <summary>Creates the exception for the call Excel rejected.</summary>
    public PrecedentsUnavailableException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    /// <summary>
    /// True for the HRESULTs of a call Excel or COM rejected because Excel is busy (<c>0x800AC472</c>,
    /// <c>RPC_E_CALL_REJECTED</c>, <c>RPC_E_SERVERCALL_RETRYLATER</c>): worth retrying, not an answer.
    /// </summary>
    public static bool IsExcelBusy(int hresult) => hresult == VbaIgnore || hresult == CallRejected || hresult == RetryLater;
}
