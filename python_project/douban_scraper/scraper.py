# -*- coding: utf-8 -*-
"""
豆瓣读书 TOP250 图书爬虫

本模块旨在从豆瓣读书 TOP250 排行榜 (book.douban.com/top250) 爬取数据。
它为每本书收集各种信息，包括：
- 书名
- 作者
- 评分
-评价人数
- 出版年份
- 出版社
- 简短引言/描述
- 书籍页面的链接

爬取的数据随后保存到名为 'douban_top250.xlsx' 的 Excel 文件中，
该文件将位于 'python_project' 目录下。

此脚本主要使用以下库：
- requests: 用于发出 HTTP GET 请求以获取网页内容。
- BeautifulSoup (来自 bs4): 用于解析页面的 HTML 结构。
- openpyxl: 用于创建 .xlsx Excel 文件并向其写入数据。
"""
import requests
from bs4 import BeautifulSoup
import openpyxl
import time

# Function to fetch HTML content
def fetch_page_html(url):
    """Fetches HTML content from a given URL."""
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
    }
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()  # Raise an exception for bad status codes
        return response.text
    except requests.exceptions.RequestException as e:
        print(f"Error fetching URL {url}: {e}")
        return None

# Function to parse HTML and extract book information
def parse_book_info(html_content):
    """Parses HTML to extract book information."""
    soup = BeautifulSoup(html_content, 'html.parser')
    books = []
    
    for item in soup.find_all('tr', class_='item'):
        book_data = {}
        
        # Title and link
        title_tag = item.find('a', title=True)
        if title_tag:
            book_data['Title'] = title_tag['title']
            book_data['Link'] = title_tag['href']
        else:
            book_data['Title'] = 'N/A'
            book_data['Link'] = 'N/A'

        # Author, Publisher, Publication Year
        pub_info_tag = item.find('p', class_='pl')
        if pub_info_tag:
            pub_info_text = pub_info_tag.get_text()
            parts = pub_info_text.split('/')
            book_data['Author'] = parts[0].strip() if len(parts) > 0 else 'N/A'
            
            # Attempt to extract Publisher, Year based on typical patterns
            if len(parts) > 1:
                # Publisher is often the last or second to last part before year
                # Year is often a 4-digit number
                year_publisher_info = "/".join(parts[1:]) # Re-join parts that might contain publisher and year
                
                # Try to find year first
                import re
                year_match = re.search(r'\b(\d{4})\b', year_publisher_info)
                if year_match:
                    book_data['Publication Year'] = year_match.group(1)
                    # Assume publisher is text before the year
                    publisher_text = year_publisher_info[:year_match.start()].strip(' /')
                    book_data['Publisher'] = publisher_text if publisher_text else 'N/A'
                else:
                    book_data['Publication Year'] = 'N/A'
                    # If no year, assume the remaining part is publisher
                    book_data['Publisher'] = year_publisher_info.strip(' /') if year_publisher_info else 'N/A'

            else: # If only author is present in parts
                book_data['Publisher'] = 'N/A'
                book_data['Publication Year'] = 'N/A'
        else:
            book_data['Author'] = 'N/A'
            book_data['Publisher'] = 'N/A'
            book_data['Publication Year'] = 'N/A'

        # Rating
        rating_tag = item.find('span', class_='rating_nums')
        if rating_tag:
            book_data['Rating'] = rating_tag.get_text()
        else:
            book_data['Rating'] = 'N/A'

        # Number of raters
        raters_tag = item.find('span', class_='pl')
        if raters_tag and "人评价" in raters_tag.get_text():
             # Extract number from (xxxx人评价)
            raters_text = raters_tag.get_text()
            num_raters_match = re.search(r'\((\d+)人评价\)', raters_text)
            if num_raters_match:
                book_data['Number of Raters'] = num_raters_match.group(1)
            else: # Fallback for cases like (少于10人评价) or other formats
                book_data['Number of Raters'] = raters_text.strip().replace('(','').replace(')','')
        else: # If not found in the typical span, search more broadly or mark N/A
            all_spans = item.find_all('span', class_='pl')
            found_raters = False
            for span in all_spans:
                if "人评价" in span.get_text():
                    raters_text = span.get_text()
                    num_raters_match = re.search(r'\((\d+)人评价\)', raters_text)
                    if num_raters_match:
                        book_data['Number of Raters'] = num_raters_match.group(1)
                    else:
                        book_data['Number of Raters'] = raters_text.strip().replace('(','').replace(')','')
                    found_raters = True
                    break
            if not found_raters:
                 book_data['Number of Raters'] = 'N/A'


        # Short description/quote
        quote_tag = item.find('span', class_='inq')
        if quote_tag:
            book_data['Quote'] = quote_tag.get_text()
        else:
            book_data['Quote'] = 'N/A'
            
        books.append(book_data)
    return books

# Main function to orchestrate scraping
def scrape_douban_top250():
    """Scrapes all pages of Douban TOP250 books."""
    base_url = "https://book.douban.com/top250?start="
    all_books = []
    # Douban TOP250 has 10 pages (0 to 9, 25 books per page)
    for i in range(10): 
        page_num = i * 25
        page_url = base_url + str(page_num)
        print(f"Scraping page {i+1}: {page_url}")
        
        html = fetch_page_html(page_url)
        if html:
            books_on_page = parse_book_info(html)
            if books_on_page:
                 all_books.extend(books_on_page)
                 print(f"Found {len(books_on_page)} books on page {i+1}.")
            else:
                print(f"No books found on page {i+1}. The page structure might have changed.")
            # Add a small delay to avoid overwhelming the server
            time.sleep(1) 
        else:
            print(f"Skipping page {i+1} due to fetch error.")
            
    return all_books

# Function to save data to Excel
def save_to_excel(books_data, filename="douban_top250.xlsx"):
    """Saves book data to an Excel file."""
    if not books_data:
        print("No data to save.")
        return

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    
    # Define header row
    headers = ["Title", "Author", "Rating", "Number of Raters", "Publication Year", "Publisher", "Quote", "Link"]
    sheet.append(headers)
    
    # Add data rows
    for book in books_data:
        row = [
            book.get("Title", "N/A"),
            book.get("Author", "N/A"),
            book.get("Rating", "N/A"),
            book.get("Number of Raters", "N/A"),
            book.get("Publication Year", "N/A"),
            book.get("Publisher", "N/A"),
            book.get("Quote", "N/A"),
            book.get("Link", "N/A")
        ]
        sheet.append(row)
    
    try:
        # Save in the python_project directory
        filepath = f"../{filename}" # Go up one level from douban_scraper to python_project
        workbook.save(filepath)
        print(f"Data saved to {filepath}")
    except Exception as e:
        print(f"Error saving Excel file: {e}")

# Main execution block
if __name__ == "__main__":
    print("Starting Douban Top 250 Book Scraper...")
    scraped_books = scrape_douban_top250()
    
    if scraped_books:
        print(f"\nSuccessfully scraped {len(scraped_books)} books in total.")
        save_to_excel(scraped_books, "douban_top250.xlsx")
    else:
        print("\nNo books were scraped. Exiting.")

    print("Scraping process finished.")
