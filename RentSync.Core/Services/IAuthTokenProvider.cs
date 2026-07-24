namespace RentSync.Core.Services;

/// <summary>
/// Supplies the bearer token for API calls.
///
/// Core stays platform-neutral: the VSTO layer registers a Windows Credential
/// Manager implementation, tests register a fake. This is the seam that makes
/// the REST client unit-testable without Windows.
/// </summary>
public interface IAuthTokenProvider
{
    /// <summary>Returns the current token, or null/empty for anonymous demo mode.</summary>
    Task<string?> GetTokenAsync(CancellationToken cancellationToken = default);
}

/// <summary>Anonymous provider used by the demo endpoint and as a safe default.</summary>
public sealed class AnonymousTokenProvider : IAuthTokenProvider
{
    public Task<string?> GetTokenAsync(CancellationToken cancellationToken = default) =>
        Task.FromResult<string?>(null);
}
