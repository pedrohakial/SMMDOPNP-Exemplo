using System.Runtime.InteropServices;

namespace Core.MemoryBroker.Models
{
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct Tick
    {
        public int TickId;
        public double Price;
        public long TimestampTicks;
    }
}
