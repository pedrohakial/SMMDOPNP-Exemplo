# Excel High-Frequency Data Pipeline Benchmark

This monorepo demonstrates the extreme performance differences between modern **Excel-DNA (via native XLL C-API + RTD)** and legacy **VSTO (via COM Interop)** when handling high-frequency data pipelines pushed directly into a Memory-Mapped File (RAM cache).

## Architecture

The system consists of the following projects:

1. **`Core.MemoryBroker` (.NET Standard 2.0):**
   - A shared library managing a cross-process, OS-level `MemoryMappedFile`.
   - Utilizes `Span<T>` and strict, unmanaged C# structs (`[StructLayout(LayoutKind.Sequential)]`) for zero-allocation memory reads/writes.
   - Synchronized using a system-level named `Mutex`.

2. **`Simulator.DataPipeline` (.NET 8 Console App):**
   - The data ingestion worker.
   - Pushes thousands of variable price updates per second into the Memory-Mapped File to simulate high-frequency tick data.

3. **`Tests.Benchmarks` (.NET 8 Console App):**
   - An isolated `BenchmarkDotNet` suite.
   - Rigorously measures the ingestion latency of writing 100,000 updates to the RAM cache without the noise of the Excel GUI.

4. **`AddIn.ExcelDNA` (.NET 8 Class Library):**
   - The modern add-in using `ExcelDna.AddIn`.
   - Exposes an RTD Server to push live memory updates back to the spreadsheet completely asynchronously.
   - Provides a `=DNA_READ(TickId)` UDF.

5. **`AddIn.VSTO` (.NET Framework 4.8 Class Library):**
   - The legacy add-in using standard VSTO boilerplate.
   - Implements a background Task/Timer that polls the MMF and writes directly to cells using `Worksheet.Range.Value2`.

---

## How to Build and Run

### Prerequisites
- Windows OS with Microsoft Excel installed.
- Visual Studio 2022 (with Office/SharePoint development workload for VSTO).
- .NET 8 SDK & .NET Framework 4.8 Developer Pack.

### Building
Open `ExcelBenchmark.sln` in Visual Studio and build the entire solution.

### Step-by-Step Instructions for the "Live Dry Run"

1. **Start the Simulator:**
   - Open a command prompt, navigate to `Simulator.DataPipeline` and run: `dotnet run -c Release`.
   - Keep this console window open. It will continuously push high-frequency mock data to the MMF and log ingestion speeds.

2. **Run the VSTO Add-In:**
   - Set `AddIn.VSTO` as the startup project in Visual Studio.
   - Hit **F5** to start debugging. This will launch Excel and load the VSTO Add-In.
   - Open a blank workbook.
   - Watch the debug console in Visual Studio or attach a debugger. The Add-In is hardcoded to poll ticks 1 to 10 and update cells `A1` to `A10`.
   - *Observation:* Notice how Excel drops frames, freezes, or entirely locks up if you try to type in other cells while data is streaming in.

3. **Run the Excel-DNA Add-In:**
   - Close the VSTO instance of Excel.
   - Set `AddIn.ExcelDNA` as the startup project.
   - Ensure the `launchSettings.json` is correctly pointing to your local `EXCEL.EXE`.
   - Hit **F5** to start debugging.
   - Open a blank workbook and enter the formula `=DNA_READ(1)` in any cell. Drag it down to row 10.
   - *Observation:* Notice how the prices update flawlessly. You can click around, type in other cells, and interact with the UI without any freezing. Check the debug logs to see the incredibly low latency (often under 1ms).

4. **Run the Micro-Benchmarks:**
   - Stop Excel and the Simulator.
   - Navigate to `Tests.Benchmarks`.
   - Run: `dotnet run -c Release`.
   - BenchmarkDotNet will measure the raw ingestion speed of the Memory-Mapped File cache.

---

## Load Test Results: Why VSTO Fails and Excel-DNA Wins

The main objective of this repository is to demonstrate *architectural bottlenecks*.

### The VSTO Bottleneck (COM Interop)
The VSTO implementation polls the RAM cache and invokes `Excel.Range.Value2`. Each time it does this, it must cross the dreaded **Managed-to-Unmanaged COM Interop boundary**.
1. COM calls run on the main Excel UI thread.
2. If your background thread fires 1,000 updates a second, you are queuing 1,000 COM marshaling requests onto the UI thread.
3. Excel's event loop becomes saturated processing COM instructions, leaving zero CPU cycles to handle user input (mouse clicks, typing, scrolling). **Result: The UI freezes.**

### The Excel-DNA Advantage (XLL C-API + RTD)
Excel-DNA utilizes the native **Excel C API (XLL)** which fundamentally bypasses COM entirely.
Furthermore, it utilizes Excel's native **RTD (Real-Time Data)** mechanism:
1. The background thread safely polls the memory-mapped file independently.
2. When a value changes, it simply flags a topic as "dirty" via RTD.
3. Excel groups these updates and recalculates the dependency tree at its own safe cadence.
4. Because memory reads are zero-allocation (`Span<T>` and unmanaged structs), the GC (Garbage Collector) never pauses the thread.
**Result: Fluid UI, zero freezes, and nanosecond memory reads.**
