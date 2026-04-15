using System;
using System.IO.MemoryMappedFiles;
using System.Runtime.InteropServices;
using System.Threading;
using Core.MemoryBroker.Models;

namespace Core.MemoryBroker.Services
{
    public class MemoryMappedCache : IDisposable
    {
        public const string MapName = "ExcelBenchmark_TickDataCache";
        public const string MutexName = "Global\\ExcelBenchmark_TickDataMutex";
        public const int MaxTicks = 10000;

        private readonly int _tickSize;
        private readonly long _capacity;

        private MemoryMappedFile _mmf;
        private MemoryMappedViewAccessor _accessor;
        private Mutex _mutex;
        private bool _disposed;

        public MemoryMappedCache()
        {
            _tickSize = Marshal.SizeOf(typeof(Tick));
            _capacity = _tickSize * MaxTicks;

            // Create or open the memory mapped file
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                _mmf = MemoryMappedFile.CreateOrOpen(MapName, _capacity, MemoryMappedFileAccess.ReadWrite);
            }
            else
            {
                // Linux fallback for sandbox testing. Named maps are entirely not supported unless name is null.
                string tempFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), MapName);
                _mmf = MemoryMappedFile.CreateFromFile(
                    tempFile,
                    System.IO.FileMode.OpenOrCreate,
                    null,
                    _capacity,
                    MemoryMappedFileAccess.ReadWrite);
            }

            _accessor = _mmf.CreateViewAccessor(0, _capacity);

            // Allow multiple processes to sync
            bool createdNew;
            _mutex = new Mutex(false, MutexName, out createdNew);
        }

        public void WriteTick(Tick tick)
        {
            if (tick.TickId < 0 || tick.TickId >= MaxTicks)
                throw new ArgumentOutOfRangeException(nameof(tick.TickId), "TickId must be between 0 and 9999");

            long offset = tick.TickId * _tickSize;

            _mutex.WaitOne();
            try
            {
                _accessor.Write(offset, ref tick);
            }
            finally
            {
                _mutex.ReleaseMutex();
            }
        }

        public Tick ReadTick(int tickId)
        {
            if (tickId < 0 || tickId >= MaxTicks)
                throw new ArgumentOutOfRangeException(nameof(tickId), "TickId must be between 0 and 9999");

            long offset = tickId * _tickSize;
            Tick tick;

            _mutex.WaitOne();
            try
            {
                _accessor.Read(offset, out tick);
            }
            finally
            {
                _mutex.ReleaseMutex();
            }

            return tick;
        }

        public void WriteTicks(Span<Tick> ticks)
        {
            _mutex.WaitOne();
            try
            {
                for (int i = 0; i < ticks.Length; i++)
                {
                    ref Tick tick = ref ticks[i];
                    if (tick.TickId >= 0 && tick.TickId < MaxTicks)
                    {
                        long offset = tick.TickId * _tickSize;
                        _accessor.Write(offset, ref tick);
                    }
                }
            }
            finally
            {
                _mutex.ReleaseMutex();
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _accessor?.Dispose();
                _mmf?.Dispose();
                _mutex?.Dispose();
                _disposed = true;
            }
        }
    }
}
