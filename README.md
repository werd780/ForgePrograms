# ForgePrograms
This is a program created to assist forges in local equipment management and tracking. 

The program uses local copies provided by J4 or local trackers matching the provided template format (Blank Template.xlsx located in Examples) for signing in and out equipment, managing individual level SHRs, and generating 1150s. 

![FP GUI Screenshot](https://github.com/user-attachments/assets/bfebb532-f8da-470c-919c-09b2358a2a79)

IMPORT - For initial setup, use your J4 HR and the the Blank Template provided and run the import function from the launcher. Then use your current local tracker to fill in the remaining SHR information and statuses. SHR information goes into the Location column (See Example Site HR.xlsx located in Examples).

INVENTORY - The Inventory function updates the "Last Scanned" and "Last Verified" columns for existing items on Site HRs. Additionally, this function can be used to track off-the-books or new equipment not yet added to the HR. By scanning or typing the asset tag, the function creates a new entry in the Excel tracker. Off-the-books items will generate a 50### series ID. If a duplicate asset tag is identified the function creates a new entry with a 51### series ID.

1150 GENNER - This function generates 1150s based off existing SHRs under column I "Loc". SHRs in this column must follow the format of "SHR: Name" for the function to properly identify them. 

1. Select the HR excel.
2. Users will be prompted to select a SHR. To be able to select a SHR, at least one entry in under "Loc" must exist. Users must manually enter the information in the column for at least one item.

![FP Genner SHR Select](https://github.com/user-attachments/assets/5190560b-4588-4223-8e1d-ab7d36c3516b)

3. Users will then be prompted to "Save Filled DD 1150 As." 
4. Finally, Users will be prompted to input the 1150 issurer and reciepient and select the transaction type from the radial selections.

![FP Genner FROM_TO_TRANSATION](https://github.com/user-attachments/assets/c4aadd0d-04b9-4164-bc38-d0115bd37b1c)

The completed 1150 will be saved under the name and location selected in step 3.

SCAN OUT - This function allows HR holders to add additional assets to a SHR. The dropdown will populate with all known SHRs pulled from column I "Loc" but you can also type in a new SHR name. The program will automatically add the "SHR: " text if you do not type it in. Users select or type the SHR they are modifying. Users can swap between SHRs if they have multiple updates to make. New items or Off-the-books items can also be added and will follow the 5#### series numbering listed under the Inventory section. After scanning/entering an asset, columns I, X, Z, and AA will be updated.

![FP Scan Out](https://github.com/user-attachments/assets/6f495a72-2b20-488a-ab76-48eb86bb2121)

SCAN IN - This function mirrors the functionality of the SCAN OUT function in reverse. This function identifies the physical locations at local Forges where assets will be stored. You can choose an existing location from the dop down or type in a new location. Locations do not need to follow the naming convention identified under the INVENTORY and SCAN OUT sections. Users can swap between locations if they have multiple updates to make. New items or Off-the-books items can also be added and will follow the 5#### series numbering listed under the Inventory section. After scanning/entering an asset, columns I, X, Z, and AA will be updated.
