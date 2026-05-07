# Excel Auto-Save Controller: Keeping Your Workflow Safe and Sound

The last thing you want to do is lose your work in the middle of a complex spreadsheet. **Excel Auto-Save Controller** is a simple, easy tool to bridge the gap between manual saving and Excel's built-in recovery features. It gives you a user-controlled, intentional safety net for your most important workbooks.

### Why It Matters
Excel does have “AutoRecover,” but it doesn’t always kick in when you need it most—like when you’re knee-deep in data entry or juggling a bunch of large files. This tool gives you back control, so you decide exactly when and how often your work is committed to disk.

### Main Features
*   **Precision Timing:** Set your own save intervals in seconds. From a heartbeat save every minute, to a broader check every half hour, the choice is yours.
*   **Multiple File Management**: Queue multiple workbooks (XLSX, XLSM, XLSB, CSV) and keep them all in sync with one "Start" button.
*   **Smart Resilience:** Built-in patience logic. If Excel is busy (e.g. you are typing in a cell, or a menu is open) then the controller will not crash, it will wait for a clear moment to do a clean save without stopping your flow.
*   **Unobtrusive Presence:** Minimize to system tray and let it protect your progress in the background. Keep your workspace clutter free.
**Real-Time Status:** Get a clear status dashboard that tells you exactly when your last save took place, and warns you if a file was momentarily skipped because Excel was “busy.”

#### Who Is This For?
**Data Analysts** working on large volatile data sets.
*   **Financial Professionals** in fast paced environments where every entry counts.
**Project Managers** juggling lots of moving parts across multiple spreadsheets.

### First steps
1. You will need to have Python installed.
2. Install dependencies: pip install xlwings pystray Pillow
3. Run the application: python Excel_Auto_Save_Controller.py

**Never again look for the 'Undo' button on a lost hour of work. Let the Excel Auto-Save Controller do the clicking so you can do the content.