# Digital Library Sheets Automation

Transform your Google Sheets into a powerful, automated digital library management system. This project utilizes advanced Google Apps Script to streamline cataloging, tracking, and managing library resources, effectively turning a standard spreadsheet into a dynamic database application.

## 🚀 Key Features

* **Event-Driven Data Routing (`AutoMove.gs`):** Features custom `onEdit` triggers that automatically transfer records between sheets based on status updates (e.g., dynamically moving a book from the "Wishlist" to the "Reading list" when marked as "Started").
* **Advanced Data Validation:** Overcomes native Google Sheets limitations by implementing custom multi-select logic for dropdown menus, allowing for complex tagging (such as in the DNF Log).
* **Smart External Data Sync (`Good reads Archive.gs`):** Features a robust synchronization script that parses Goodreads export data and intelligently routes books. It includes string-normalization for strict deduplication, automated status mapping (routing "to-read" to the Wishlist and "read/dnf" to the Tracker), and a custom gap-free appending algorithm to maintain database integrity.
* **Custom User Interface (`Sidebar.html`, `RatingsGuide.html`):** Enhances the native spreadsheet environment with custom HTML/CSS sidebars and dialog boxes, creating a seamless and user-friendly graphical interface.
* **Automated Asset Management:** Utilizes utility scripts (`Downloader.gs`, `Renamer.gs`) to streamline file handling and resource management in the background.

## 🛠️ Technologies Used

* **Google Apps Script** (JavaScript)
* **HTML / CSS** (For custom UI components)
* **Google Sheets API**

## 📦 Setup and Installation

1. Create a new Google Sheet.
2. Navigate to `Extensions` > `Apps Script`.
3. Clear the default code and copy the contents of the `.gs` and `.html` files in the `src` directory into your Apps Script editor.
4. Save the project and refresh your Google Sheet to initialize the custom menus and triggers.
5. *(Optional)* Run the initial setup functions to generate the required sheet tabs (e.g., Wishlist, Reading list, My library, DNF Log).

## 🤝 Contributing
Contributions, issues, and feature requests are welcome! Feel free to check the [issues page](https://github.com/Rei0322/digital-lib-sheets/issues) if you want to contribute.

## 👤 Author
**Samira N Ray**
* GitHub: [@Rei0322](https://github.com/Rei0322)
