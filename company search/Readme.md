## Company Info Extractor ##

This program retrieves detailed, up-to-date information about multiple companies using a paid **Gemini API** (advanced pro model) and displays the results in real time via a Flask web interface.

## What Problem Does It Solve?

1. **Centralized Data Retrieval**  
   It combines AI-driven data (via Gemini) with fallback sources (like Wikipedia or hardcoded info) to gather company information in one place, saving the user from manually researching each company’s details.

2. **Real-Time Feedback**  
   The app provides a progress bar and estimated time remaining, helping users track the status of their queries when processing multiple companies at once.

3. **Business Intelligence & Lead Generation**  
   By collecting crucial data—industry, employee size, revenue, contacts, etc.—the program can aid market research, sales, or competitive analysis.

## High-Level Overview

1. **Flask Application**  
   - Runs a web server where users can input a list of companies.  
   - Displays the extraction progress and final results in a responsive interface.

2. **Gemini API Integration**  
   - Uses a **paid pro model** (`gemini-2.0-pro`) for advanced data retrieval.  
   - The prompt instructs the AI to provide verified data in strict JSON format.  
   - If the AI or data is unavailable, the app returns placeholders or “N/A” to avoid guesses.

3. **Concurrent Processing & Real-Time Updates**  
   - A background thread processes each company concurrently (via `ThreadPoolExecutor`).  
   - The front end polls a `/progress/<job_id>` endpoint every second, updating the UI as each company’s info is ready.

4. **Fallback & Error Handling**  
   - If the model can’t find data or hits rate limits, the app retries or provides fallback info (e.g., from Wikipedia or hardcoded entries).  
   - The user sees partial results as they become available, and any errors are reported clearly.

## Code Structure

- **`CompanyInfoExtractor` class**  
  - Initializes the Gemini API with a given key and model name (`gemini-2.0-pro`).  
  - `extract_company_info(company_name)`:  
    - Checks for hardcoded info or Wikipedia context.  
    - Sends a carefully crafted prompt to Gemini.  
    - Parses the JSON response or handles errors.  
  - `batch_extract(companies, ...)`:  
    - Processes multiple companies in parallel, updating real-time progress via a callback.

- **Flask Routes**  
  - **`/`**: Renders the main HTML page with a form for company names.  
  - **`/start`**: Accepts a POST request of company names, starts the extraction job in a separate thread, and returns a `job_id`.  
  - **`/progress/<job_id>`**: Returns the real-time status (how many companies are done, estimated time, final results if complete).

- **Threading & Real-Time Progress**  
  - Each job runs in a separate thread to avoid blocking.  
  - The front end polls the progress endpoint, updating a progress bar and displaying results incrementally.

## How to Use

1. **Install Dependencies**  
   pip install flask google-generativeai wikipedia


2. **Set Your Gemini API Key**

In the code, replace the placeholder with your actual paid pro key:
python
Copy
EXTRACTOR = CompanyInfoExtractor("YOUR_GEMINI_PRO_KEY_HERE")
Run the Flask App
Copy
python app.py
Then open http://127.0.0.1:5000/ in your browser.

3. **Enter Company Names**

Provide one or more company names (comma or newline separated).

Click Start Extraction to see real-time progress and final data.

Copy
