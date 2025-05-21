# Python Project Collection

## Overall Project Description
This repository houses various Python modules, each designed for a specific function or task. The aim is to provide a collection of useful scripts and tools developed in Python.

## Modules

### Available Modules

#### Douban Book Scraper

*   **Description:** This module scrapes data for the TOP250 books from Douban Reading (读书 - book.douban.com).
*   **Functionality:** It fetches detailed information for each book, including its title, author, rating, number of raters, publication year, publisher, a short quote/description, and a direct link to the book's Douban page.
*   **Output:** The scraped data is saved in an Excel file named `douban_top250.xlsx`, located in the main `python_project` directory.
*   **How to Run:**
    1.  Ensure you have Python installed, along with the necessary libraries: `requests`, `beautifulsoup4`, and `openpyxl`. If not, you can install them using pip:
        ```bash
        pip install requests beautifulsoup4 openpyxl
        ```
    2.  Navigate to the scraper's directory:
        ```bash
        cd python_project/douban_scraper/
        ```
    3.  Run the scraper script:
        ```bash
        python scraper.py
        ```
        The script will print its progress and inform you once the data is saved.

## Future Development
This project is under continuous development. More Python modules and useful scripts will be added in the future. Stay tuned!
