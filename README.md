# Excel Auto-Save Controller: Peace of Mind for Your Workflow

In the middle of a complex spreadsheet, the last thing you should worry about is losing progress. The **Excel Auto-Save Controller** is a lightweight, intuitive utility designed to bridge the gap between manual saving and Excel's native recovery features. It provides a deliberate, user-controlled safety net for your most critical workbooks.

### Why It Matters
While Excel has "AutoRecover," it doesn't always trigger when you need it most—especially during intense data entry or when working across multiple heavy files. This tool puts the power back in your hands, allowing you to define exactly when and how often your work is committed to disk.

### Key Features
*   **Precision Timing:** Set custom save intervals down to the second. Whether you need a heartbeat save every minute or a broader check every half hour, the choice is yours.
*   **Multi-File Management:** Queue up multiple workbooks (XLSX, XLSM, XLSB, CSV) and keep them all synchronized with a single "Start" button.
*   **Smart Resilience:** Built-in "patience" logic. If Excel is busy (like when you're actively typing in a cell or a menu is open), the controller won't crash; it gracefully waits for a clear moment to ensure a clean save without interrupting your flow.
*   **Unobtrusive Presence:** Minimize the application to the system tray to keep your workspace clutter-free while it continues to guard your progress in the background.
*   **Real-Time Status:** Stay informed with a clear status dashboard showing exactly when your last save occurred and alerting you if a file was temporarily skipped due to Excel being "busy."

### Who Is This For?
*   **Data Analysts** managing large, volatile datasets.
*   **Financial Professionals** working in high-stakes environments where every entry counts.
*   **Project Managers** coordinating multiple moving parts across various spreadsheets.

### Getting Started
1. Ensure you have Python installed.
2. Install dependencies: `pip install xlwings pystray Pillow`
3. Run the application: `python Excel_Auto_Save_Controller.py`

**Never look for the "Undo" button on a lost hour of work again. Let the Excel Auto-Save Controller handle the clicks so you can focus on the content.**
