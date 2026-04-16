using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Core.MemoryBroker.Models;
using Core.MemoryBroker.Services;

namespace AddIn.VSTO
{
    public partial class ThisAddIn
    {
        private MemoryMappedCache _cache;
        private CancellationTokenSource _cts;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _cache = new MemoryMappedCache();
            _cts = new CancellationTokenSource();

            // Deliberately run a background task that polls the cache and updates UI via COM Interop
            Task.Run(() => PollingLoop(_cts.Token), _cts.Token);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            _cts?.Cancel();
            _cache?.Dispose();
        }

        private void PollingLoop(CancellationToken token)
        {
            // Watch specific ticks, e.g., Tick 1 to 10
            int[] watchedTicks = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };

            while (!token.IsCancellationRequested)
            {
                try
                {
                    Excel.Worksheet activeSheet = this.Application.ActiveSheet;
                    if (activeSheet != null)
                    {
                        foreach (int tickId in watchedTicks)
                        {
                            Tick tick = _cache.ReadTick(tickId);

                            long latencyTicks = DateTime.UtcNow.Ticks - tick.TimestampTicks;
                            double latencyMs = TimeSpan.FromTicks(latencyTicks).TotalMilliseconds;

                            Debug.WriteLine($"[VSTO] Latency for Tick {tickId}: {latencyMs:F2} ms");

                            // PERFORMANCE BOTTLENECK: The COM interop call here is slow and will eventually block the Excel UI
                            // as the frequency increases.
                            string cellAddress = $"A{tickId}";
                            // In real VSTO activeSheet.Range works. We mock it with get_Range
                            Excel.Range range = activeSheet.get_Range(cellAddress);
                            range.Value2 = $"Tick={tick.TickId} Price={tick.Price:F4}";
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[VSTO] Polling Error: {ex.Message}");
                }

                Thread.Sleep(100); // 100ms poll interval
            }
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
