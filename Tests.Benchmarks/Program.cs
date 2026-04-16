using System;
using System.Diagnostics;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using Core.MemoryBroker.Models;
using Core.MemoryBroker.Services;

namespace Tests.Benchmarks
{
    [MemoryDiagnoser]
    public class CacheIngestionBenchmark
    {
        private MemoryMappedCache _cache;
        private Tick[] _batchData;

        [GlobalSetup]
        public void Setup()
        {
            _cache = new MemoryMappedCache();
            _batchData = new Tick[100];
            var random = new Random(42);

            for (int i = 0; i < _batchData.Length; i++)
            {
                _batchData[i] = new Tick
                {
                    TickId = random.Next(0, MemoryMappedCache.MaxTicks),
                    Price = 100.0 + random.NextDouble() * 10,
                    TimestampTicks = DateTime.UtcNow.Ticks
                };
            }
        }

        [GlobalCleanup]
        public void Cleanup()
        {
            _cache?.Dispose();
        }

        [Benchmark]
        public void Write100000Updates()
        {
            // We write _batchData (100 ticks) 1000 times = 100,000 updates
            Span<Tick> batchSpan = _batchData.AsSpan();
            for (int i = 0; i < 1000; i++)
            {
                // To avoid constant re-reading of UtcNow in tight benchmark loops
                // while isolating just the memory-mapped writing performance,
                // we'll update the timestamp once per batch.
                long ticks = DateTime.UtcNow.Ticks;
                for (int j = 0; j < batchSpan.Length; j++)
                {
                    batchSpan[j].TimestampTicks = ticks;
                }

                _cache.WriteTicks(batchSpan);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            var summary = BenchmarkRunner.Run<CacheIngestionBenchmark>();
        }
    }
}
