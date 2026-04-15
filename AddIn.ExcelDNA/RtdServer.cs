using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Threading;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;
using Core.MemoryBroker.Models;
using Core.MemoryBroker.Services;

namespace AddIn.ExcelDNA
{
    public class RtdServer : ExcelRtdServer
    {
        private MemoryMappedCache _cache;
        private Timer _timer;
        private ConcurrentDictionary<int, Topic> _topics = new ConcurrentDictionary<int, Topic>();

        protected override bool ServerStart()
        {
            try
            {
                _cache = new MemoryMappedCache();
                // Start a background timer to poll for updates and push them to registered topics.
                // Doing this at 100ms or even lower frequency to show how RTD handles pushing updates
                _timer = new Timer(PollUpdates, null, 0, 100);
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error starting RTD server: {ex.Message}");
                return false;
            }
        }

        protected override void ServerTerminate()
        {
            _timer?.Dispose();
            _cache?.Dispose();
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            if (topicInfo.Count > 0 && int.TryParse(topicInfo[0], out int tickId))
            {
                _topics[tickId] = topic;

                // Provide initial value
                try
                {
                    Tick tick = _cache.ReadTick(tickId);
                    return FormatOutput(tick);
                }
                catch
                {
                    return ExcelError.ExcelErrorNA;
                }
            }

            return ExcelError.ExcelErrorValue;
        }

        protected override void DisconnectData(Topic topic)
        {
            foreach (var kvp in _topics)
            {
                if (kvp.Value == topic)
                {
                    _topics.TryRemove(kvp.Key, out _);
                    break;
                }
            }
        }

        private void PollUpdates(object state)
        {
            foreach (var kvp in _topics)
            {
                int tickId = kvp.Key;
                Topic topic = kvp.Value;

                try
                {
                    Tick tick = _cache.ReadTick(tickId);

                    // Here we log the latency difference between the tick ingestion timestamp and Excel RTD update.
                    // This logs the time taken from RAM directly to the point where RTD asks Excel to recalculate.
                    long latencyTicks = DateTime.UtcNow.Ticks - tick.TimestampTicks;
                    double latencyMs = TimeSpan.FromTicks(latencyTicks).TotalMilliseconds;

                    // Send out to debug or console
                    Debug.WriteLine($"[Excel-DNA] Latency for Tick {tickId}: {latencyMs:F2} ms");

                    topic.UpdateValue(FormatOutput(tick));
                }
                catch
                {
                    // Ignore errors during poll
                }
            }
        }

        private string FormatOutput(Tick tick)
        {
            return $"Tick={tick.TickId} Price={tick.Price:F4}";
        }
    }

    public static class Functions
    {
        [ExcelFunction(Description = "Reads tick data from the Memory-Mapped File directly with zero-allocation")]
        public static object DNA_READ(int tickId)
        {
            return XlCall.RTD("AddIn.ExcelDNA.RtdServer", null, tickId.ToString());
        }
    }
}
