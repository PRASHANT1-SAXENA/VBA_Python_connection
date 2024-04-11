# VBA_Python_connection

    1. Create a Simple XLSX File: Begin by creating a basic Excel file (.xlsx) using Excel or any spreadsheet software. You can add some sample data or leave it blank for now.

    2. Save the File as a Macro-Enabled Workbook (.xlsm): After creating the .xlsx file, save it as a macro-enabled workbook (.xlsm). Most spreadsheet software allows you to save files in different formats. Ensure to select the ".xlsm" format when saving.

    3. Install Office RibbonX Editor: Download and install the Office RibbonX Editor, which is a tool specifically designed for editing custom ribbon interfaces in Microsoft Office applications.

    4. Open the Excel File with RibbonX Editor: Launch the RibbonX Editor and open the Excel file you created in step 2. This will allow you to modify the ribbon interface.

    5. Create XML Script for Custom Ribbon: In the RibbonX Editor, you'll find options to create a new custom ribbon interface or modify an existing one. Write the XML script for your custom ribbon interface. This script defines the layout and functionality of the ribbon tabs, groups, and buttons.

    6. Replace the Sample Code: If there's any sample code provided by the RibbonX Editor, delete it. Paste your custom XML script into the editor.

    7. Save the XML File: Once you've finalized the custom ribbon interface script, save it as an XML file. This XML file will contain the instructions for how the ribbon interface should appear and function in your Excel workbook.

for example which i used in this file 'xml code'
#######

<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon startFromScratch="false"> 
    <tabs> 
      <tab id="KO" label="KO"> 
        <group id="login" label="Authentication"> 
             <button id="sso_login" imageMso="GroupAdminister" size="large" 
                      label="Login" 
                      screentip="login on KO LRP PLatform" 
                      onAction="open_login"/> 
                <button id="sso_logout" imageMso="PrintPreviewClose" size="large" 
                      label="LogOut" 
                      screentip="..." 
                      onAction="open_login"/> 
        </group>
            
        <group id="Control_001" label="MRP-N"> 
             <button id="b_mrpn" image="mrpn" size="large" 
                      label="Support MRPN" 
                      screentip="Ask for support of our team"
                      onAction="open_hyperlink_mrpn"/> 
             <button id="b_coca" image="ko3" size="large" 
                      label="Support KO" 
                      screentip="Contact KO Technical Team" 
                      onAction="open_hyperlink_ko"/> 
        </group>
            
        <group id="Control_002" label="SQL"> 
            <button id="b004" image="table" size="large" 
                  label="SQL Data" 
                  getEnabled="GetEnabled" 
                  screentip="This is a Tip about what this button do on MRP-N add-in" 
                  onAction="msg_open"/> 

        </group>  
            
         <group id="Control_003" label="Master Data"> 
            <button id="b4" imageMso="AccessRefreshAllLists" size="large" 
                  label="Refresh Data" 
                  screentip="This is a Tip about what this button do on MRP-N add-in" 
                  onAction="msg_open"/> 
            <button id="b61" imageMso="UpgradeWorkbook" size="large" 
                    label="Upload Data to DBMS" 
                    screentip="This is a Tip about what this button does on MRP-N add-in" 
                    onAction="RunPythonScript"/> 
            <button id="b62" imageMso="TableExcelSpreadsheetInsert" size="large" 
                  label="Get Data from DBMS" 
                  screentip="This is a Tip about what this button do on MRP-N add-in" 
                  onAction="ReadCSVFile"/> 
            <button id="b63" imageMso="TableRowsOrColumnsOrCellsInsert" size="large" 
                  label="SKU By Line" 
                  screentip="This is a Tip about what this button do on MRP-N add-in" 
                  onAction="msg_open"/> 
            <button id="b64" imageMso="DatasheetView" size="large" 
                  label="Hours By Line" 
                  screentip="This is a Tip about what this button do on MRP-N add-in" 
                  onAction="msg_open"/> 
            <button id="b65" imageMso="TableShowGridlines" size="large" 
                  label="OUBP Demand" 
                  screentip="This is a Tip about what this button do on MRP-N add-in" 
                  onAction="msg_open"/> 
            <button id="b66" imageMso="_3DStyle" size="large" 
                  label="Demand Incremental" 
                  screentip="This is a Tip about what this button do on MRP-N add-in" 
                  onAction="msg_open"/> 
            <button id="b67" imageMso="Chart3DColumnChart" size="large" 
                  label="CAPEX" 
                  screentip="This is a Tip about what this button do on MRP-N add-in" 
                  onAction="msg_open"/> 

            <button id="b7" imageMso="FunctionWizard" size="large" 
                  label="Run Simulation" 
                  screentip="This is a Tip about what this button do on MRP-N add-in" 
                  onAction="msg_open"/> 
            <button id="b1" imageMso="EqualSign" size="normal" 
                  label="IfErrorBlank" 
                  onAction="msg_open"/> 
            <button id="b3" imageMso="EqualSign" size="normal" 
                  label="IfErrorZero" 
                  onAction="msg_open"/> 
        </group> 
      </tab> 
    </tabs> 
  </ribbon> 
</customUI> 


############################


    8. Access Excel Options: Open your Excel workbook and click on the "File" tab located in the top-left corner. Then, click on the "Options" button at the bottom of the menu that appears.

    9. Customize Ribbon: In the Excel Options window, select "Customize Ribbon" from the list on the left-hand side. This allows you to customize the ribbon interface.

    10. Enable Developer Tab: Within the Customize Ribbon section, look for the "Developer" option in the list on the right-hand side. Make sure the checkbox next to "Developer" is checked. This will enable the Developer tab in the Excel ribbon.

    11. Access Developer Tab: After enabling the Developer tab, you will see it displayed in the Excel ribbon. Click on the Developer tab to access its features.

    12. Open Visual Basic: Within the Developer tab, locate and click on the "Visual Basic" button. This will open the Visual Basic for Applications (VBA) editor.

    13. Create a Module: In the VBA editor window, navigate to the "Insert" menu and choose "Module". This will create a new module where you can write VBA code.

    14. Write VBA Code: In the newly created module, write the VBA code necessary for your task. You can reference existing files or repositories for sample code, ensuring to adapt it to your specific needs.

    15. Save the VBA Module: Once you've written the VBA code, make sure to save the module within your Excel workbook.

    16. Place Python File in the Same Folder: Ensure that the Python script (main.py) is placed in the same folder as your Excel workbook (.xlsm) file. This is necessary for the VBA code to access and execute the Python script.