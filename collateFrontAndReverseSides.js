/**
 * Author: Nazmul Kazi
 * Version: 2.0.0
 * Updated: June 1, 2024
 * 
 * Collates front and reverse sides of scanned pages from two documents into one.
 * The current active document should contain the front sides and select the file
 * containing the reverse sides when prompted. Both files must contain the same number
 * of pages. After the pages are collated, the user will be prompted to save the
 * document. Since the pages are scanned using a single-sided Automatic Document
 * Feeder (ADF), the reverse sides should be in reverse order.
 */

/************************************ Functions ************************************/

// Function to collate front sides with reverse sides
var collateOddAndEvenPages = app.trustedFunction(function() {
    try {
        // Raise execution privilege
        app.beginPriv();
        
        // Prompt the user for the file with reverse sides
        var reverseSidePages = app.browseForDoc();
        
        // Check if a document is provided
        if (reverseSidePages) {
            /**
             * Both files must contain the same number of pages. `app.browseForDoc` does not
             * return any document object or metadata that contains the number of pages. One
             * way to validate the number of reverse side pages is by opening the document
             * which is not ideal because: (a) Acrobat cannot open a file in background, and
             * (b) after counting the number of pages, we cannot use the open document to
             * collate the pages. Instead, we insert all the reverse sides at the end of the
             * current active document (that contains the front sides) and count the number
             * of pages before and after inserting the reverse sides. If the number of pages
             * does not match, we remove the inserted pages and raise a page mismatch error.
             * Otherwise, we move the reverse side pages to their correct position within the
             * document to collate the pages.
             */
            
            // Count the number of front side pages
            var numFrontSidePages = this.numPages;
            
            // Insert reverse side pages at the end of current active document
            this.insertPages({
                nPage: numFrontSidePages - 1, // index of where to insert the pages
                cPath: reverseSidePages.cPath  // path to the source file
            });
            
            // Count the number of reverse side pages
            var numReverseSidePages = this.numPages - numFrontSidePages;
            
            // Check if the number of front side pages matches the number of reverse side pages.
            if (numFrontSidePages === numReverseSidePages) {
                /**
                 * Move the reverse side pages (i.e., inserted pages) to their correct position.
                 * 
                 * Since the reverse side pages are in reverse order, we start by moving the first
                 * reverse side page which is the last page in the document. After each move, the
                 * next reverse side page that needs to be moved is always the last page in the
                 * document. We do not need to move the last reverse side page as it should be the
                 * last page in the document and already in its correct position.
                 */
                for (var i = 0; i < numReverseSidePages - 1; i++) {
                    this.movePage({
                        nPage: this.numPages - 1, // index of the last page in the document
                        nAfter: i * 2 // index of the page after which the selected page should be moved
                    });
                }
                
                // Prompt the user to save the collated document
                var saveChanges = app.alert({
                    cTitle: "Save", // Title of the dialog box
                    cMsg: "Front sides are successfully collated with reverse sides. Do you want to save the updated document?",
                    nIcon: 2, // Show question icon
                    nType: 2, // Show yes/no buttons only
                });
                
                // Save the document, if requested
                if (saveChanges === 4) { // 4 corresponds to the "Yes" button
                    app.execMenuItem("Save");
                }
            } else {
                // Delete the inserted pages
                this.deletePages({
                    nStart: numFrontSidePages, // index of the first page to delete
                    nEnd: this.numPages - 1 // index of the last page to delete
                });
                
                // Reset the flag to reflect no changes are made
                this.dirty = false;
                
                // Extract the name of the file containing the reverse side pages
                var reverseSideDocFilename = reverseSidePages.cPath.match(/\/([^/]+)$/)[1];
                
                // Show an error message to user and exit
                app.alert("Both files must contain the same number of pages. The current active file \"" + this.documentFileName + "\", that should contain the front sides, has " + numFrontSidePages + " page" + (Math.abs(this.numPages) == 1 ? " only" : "s") + ". However, the file \"" + reverseSideDocFilename + "\", that should contain the reverse (i.e., back) sides, has " + numReverseSidePages + " page" + (Math.abs(numReverseSidePages) == 1 ? " only": "s") + ".");
            }
        }
    } catch (e) {
        app.alert("Encountered an unexpected error:", e.message);
    } finally {
        // End execution privilege
        app.endPriv();
    }
});

/************************************ Menu ************************************/

// Check if the new Acrobat design with the Hamburger Menu is enabled.
var newHamburgerMenu = app.listMenuItems()[0].cName === "AV2::HamburgerMenu";

// Add a new menu item
app.addMenuItem({
    // Define a unique identifier for internal use
    cName: "collateOddAndEvenPages",
    
    // Set display name for users
    cUser: "Collate with Reverse Sides",
    
    // Set parent menu:
    // - Hamburger Menu: place it at the top-level
    // - Old Menu: place it under the "Edit" menu
    cParent: newHamburgerMenu ? "AV2::HamburgerMenu" : "Edit",
    
    // Set item position in the menu:
    // - Hamburger Menu: place the item below the "Combine Files" menu item.
    // - Old Menu: place the item below the "Rotate Pages" menu item.
    nPos: newHamburgerMenu ? "CombineFilesHamburgerMenu" : "RotatePagesMenuItem",
    
    // Enable item for open and active documents only
    cEnable: "event.rc = (event.target != null);",
    
    // Set the function to execute when the menu item is selected
    cExec: "collateOddAndEvenPages();"
});