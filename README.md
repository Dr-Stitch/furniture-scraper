# Furniture Product Information Scraper

## Overview

This project is an advanced web scraping tool designed to automate the extraction of detailed product information from furniture e-commerce websites. It leverages Selenium WebDriver for browser automation and data extraction, and utilizes Python's data processing and Excel manipulation libraries to organize and store the results.

## Features

- **Automated Web Scraping:** Uses Selenium to interact with dynamic web pages, handle pop-ups, scroll, and extract hidden or dynamically loaded content.
- **Comprehensive Data Extraction:** Gathers product details such as title, price, description, highlights, images, colors, dimensions, ratings, reviews, SKU, and more.
- **Excel Integration:** Reads product URLs and categories from an input Excel file and writes the scraped data to a structured output Excel file using `pandas` and `openpyxl`.
- **Error Handling & Logging:** Implements robust error handling and logging to ensure reliability and easy debugging.
- **Data Deduplication:** Cleans and deduplicates input data before processing.
- **Scalable & Adaptable:** Designed to handle products with multiple sizes, color options, and complex page layouts.

## Skills Demonstrated

- **Python Programming:** Advanced scripting, modular code design, and use of standard libraries.
- **Web Automation & Scraping:** Proficient use of Selenium WebDriver for browser automation and dynamic content extraction.
- **Data Processing:** Efficient handling and transformation of data using `pandas`.
- **Excel Automation:** Reading, writing, and updating Excel files programmatically with `openpyxl`.
- **Error Handling & Logging:** Implementation of try/except blocks and logging for robust, production-ready code.
- **Problem Solving:** Tackling real-world challenges such as dynamic web content, pop-ups, and data normalization.

## How It Works

1. **Input Preparation:** Reads a list of product URLs and categories from `Product_List.xlsx`.
2. **Web Scraping:** For each product, launches a browser session, navigates to the product page, and extracts all relevant information.
3. **Data Storage:** Appends the extracted data to `Product Information.xlsx` in a structured format.
4. **Error Management:** Logs any issues encountered during scraping for review and troubleshooting.

## Getting Started

1. Install the required Python packages:
    ```bash
    pip install selenium pandas openpyxl
    ```
2. Ensure you have the correct WebDriver (Edge) installed and available in your PATH.
3. Prepare your `Product_List.xlsx` with product URLs and categories.
4. Run the script:
    ```bash
    python product-information-scraper.py
    ```

## Portfolio & Contact

This project demonstrates my ability to automate complex data extraction tasks, process and organize large datasets, and deliver robust, production-ready Python solutions. If you are interested in collaborating or have opportunities that match my skill set, feel free to reach out!
