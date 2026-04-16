using System;
using System.Diagnostics;
using System.Threading;
using Core.MemoryBroker.Models;
using Core.MemoryBroker.Services;

namespace Simulator.DataPipeline
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Starting Simulator.DataPipeline...");
            Console.WriteLine($"Configured for {MemoryMappedCache.MaxTicks} ticks. Press Ctrl+C to stop.");

            using var cache = new MemoryMappedCache();
            var random = new Random();
            long updatesCount = 0;
            var stopwatch = Stopwatch.StartNew();

            // Span allocated on stack for batching, let's update a batch of 100 ticks per loop iteration
            Span<Tick> batch = stackalloc Tick[100];

            try
            {
                while (true)
                {
                    long currentTicks = DateTime.UtcNow.Ticks;

                    for (int i = 0; i < batch.Length; i++)
                    {
                        // Pick a random tick id to update
                        int tickId = random.Next(0, MemoryMappedCache.MaxTicks);

                        batch[i] = new Tick
                        {
                            TickId = tickId,
                            Price = 100.0 + (random.NextDouble() * 10.0), // Random price between 100 and 110
                            TimestampTicks = currentTicks
                        };
                    }

                    // Write batch to MMF
                    cache.WriteTicks(batch);
                    updatesCount += batch.Length;

                    if (stopwatch.ElapsedMilliseconds >= 1000)
                    {
                        Console.WriteLine($"[Info] Ingested {updatesCount} updates per second.");
                        updatesCount = 0;
                        stopwatch.Restart();
                    }

                    // Simulate a small delay to avoid maxing out a single CPU core entirely while still maintaining high frequency
                    Thread.Sleep(1);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Error] {ex.Message}");
            }
        }
    }
}
