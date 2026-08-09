using System.Collections.Concurrent;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Headless;
using Avalonia.Threading;

namespace DataWizard.App.Tests;

/// <summary>
/// Boots one headless Avalonia application for the whole test run and gives tests
/// a way to execute on its UI thread.
/// </summary>
/// <remarks>
/// Avalonia can only be initialised once per process and everything it owns is
/// thread-affine, so the session runs on its own dedicated thread and work is
/// marshalled onto it.
/// </remarks>
public sealed class AvaloniaFixture : IDisposable
{
    private readonly Thread _uiThread;
    private readonly ManualResetEventSlim _ready = new(false);
    private readonly CancellationTokenSource _shutdown = new();
    private Exception? _startupFailure;

    public AvaloniaFixture()
    {
        _uiThread = new Thread(Run)
        {
            IsBackground = true,
            Name = "Avalonia headless test thread"
        };

        // A single-threaded apartment is a Windows requirement (COM, and the
        // clipboard and dialog APIs that sit on it); elsewhere the call is not
        // supported at all.
        if (OperatingSystem.IsWindows())
            _uiThread.SetApartmentState(ApartmentState.STA);

        _uiThread.Start();

        if (!_ready.Wait(TimeSpan.FromSeconds(60)))
            throw new TimeoutException("The headless Avalonia session did not start within 60 seconds.");

        if (_startupFailure is not null)
            throw new InvalidOperationException("The headless Avalonia session failed to start.", _startupFailure);
    }

    private void Run()
    {
        try
        {
            AppBuilder.Configure<App>()
                .UseHeadless(new AvaloniaHeadlessPlatformOptions { UseHeadlessDrawing = true })
                .SetupWithLifetime(new ClassicDesktopStyleApplicationLifetime());

            _ready.Set();

            Dispatcher.UIThread.MainLoop(_shutdown.Token);
        }
        catch (Exception ex)
        {
            _startupFailure = ex;
            _ready.Set();
        }
    }

    /// <summary>Runs an action on the Avalonia thread and rethrows anything it throws.</summary>
    public void Invoke(Action action) => Dispatcher.UIThread.Invoke(action);

    /// <summary>Runs a function on the Avalonia thread and returns its result.</summary>
    public T Invoke<T>(Func<T> function) => Dispatcher.UIThread.Invoke(function);

    public void Dispose()
    {
        _shutdown.Cancel();
        _uiThread.Join(TimeSpan.FromSeconds(10));
        _ready.Dispose();
        _shutdown.Dispose();
    }
}

/// <summary>Shares the single Avalonia session across every UI test class.</summary>
[CollectionDefinition(Name)]
public sealed class AvaloniaCollection : ICollectionFixture<AvaloniaFixture>
{
    public const string Name = "Avalonia";
}
