## How to Manually Add the "Convert Chart Data" Button to PowerPoint

This guide explains how to add the "Convert Chart Data" macro and its corresponding ribbon button to your PowerPoint application for use in any presentation. This process involves two main parts:
1.  **Making the Macro Globally Available:** We will save the VBA macro in a special PowerPoint startup folder.
2.  **Importing the Custom Ribbon:** We will use PowerPoint's built-in tool to import the custom ribbon layout.

### Part 1: Make the Macro Globally Available

For a custom ribbon button to work, the macro it calls must be loaded every time PowerPoint starts. The easiest way to achieve this is by placing it in a macro-enabled presentation in PowerPoint's `STARTUP` folder.

1.  **Find your PowerPoint `STARTUP` folder:**
    *   Press **Windows Key + R** to open the "Run" dialog.
    *   Type `%appdata%\Microsoft\PowerPoint\STARTUP` and press **Enter**.
    *   This will open the correct folder in File Explorer. Keep this window open.

2.  **Create a Global Macros Presentation:**
    *   Open PowerPoint and create a **new, blank presentation**.
    *   Press **Alt + F11** to open the VBA Editor.
    *   In the "Project" pane on the left, right-click `VBAProject (Presentation1)` and select **Insert > Module**.
    *   A new code window will appear. Copy the entire contents of the `macro.vba` file and paste it into this new module window.
    *   In the VBA Editor, go to **Tools > References** and ensure that **Microsoft Office 16.0 Object Library** is checked (the version number may vary).

3.  **Save the Presentation to the `STARTUP` folder:**
    *   Go to **File > Save As** in the main PowerPoint window.
    *   Navigate to the `STARTUP` folder you opened in step 1.
    *   Save the file as a **PowerPoint Macro-Enabled Presentation (*.pptm)**.
    *   You can name it something descriptive, like `GlobalMacros.pptm`.
    *   Close the presentation.

Now, whenever you start PowerPoint, this file will be loaded in the background, making the `OnConvertChartData` macro available to the ribbon.

### Part 2: Import the Custom Ribbon Layout

1.  **Open PowerPoint Options:**
    *   Go to **File > Options**.
    *   Select the **Customize Ribbon** tab on the left.

2.  **Import the Customization File:**
    *   At the bottom-right of the window, click the **Import/Export** button.
    *   From the dropdown menu, select **Import customization file**.
    *   Navigate to where you saved the `ChartTools.exportedUI` file and select it.
    *   Click **Open**.
    *   You will be asked if you want to replace all existing ribbon and Quick Access Toolbar customizations. Click **Yes**.

3.  **Finish:**
    *   Click **OK** to close the PowerPoint Options window.

You should now see a **Chart Tools** tab on your PowerPoint ribbon, located just after the "Home" tab. The "Convert Chart Data to Values" button will be inside this tab and is ready to use on any presentation.

### Troubleshooting

*   **The "Chart Tools" tab or button is missing:** Double-check that you imported the `.exportedUI` file correctly.
*   **The button is visible but does nothing when clicked:** This almost always means the macro is not loaded. Ensure that your `GlobalMacros.pptm` file (or whatever you named it) is saved in the correct `%appdata%\Microsoft\PowerPoint\STARTUP` folder and that it contains the correct macro code.
