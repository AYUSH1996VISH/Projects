import json
import time
import re
import uuid
import threading
import wikipedia
import google.generativeai as genai
from flask import Flask, render_template, request, jsonify
from concurrent.futures import ThreadPoolExecutor, as_completed

app = Flask(__name__)
app.secret_key = "your_secret_key_here"  # Replace with a secure key in production

# Global dictionary to hold job status and results
jobs = {}
MAX_COMPANIES = 50  # Maximum companies allowed per job

def extract_json_from_text(text):
    """
    Attempt to extract the first JSON object from the text using regex.
    """
    try:
        pattern = re.compile(r'\{.*\}', re.DOTALL)
        match = pattern.search(text)
        if match:
            json_str = match.group()
            return json.loads(json_str)
    except Exception as e:
        print(f"Error extracting JSON: {e}")
    return None

class CompanyInfoExtractor:
    def __init__(self, api_key):
        """
        Initialize the extractor with the Google Gemini API using gemini-2.0-flash.
        """
        genai.configure(api_key=api_key)
        self.model_name = 'gemini-2.0-flash'
        print(f"Using model: {self.model_name}")
        self.model = genai.GenerativeModel(self.model_name)
    
    def _get_hardcoded_company_info(self, company_name):
        """
        Return hardcoded info for known companies.
        """
        company_info = {
            "Google": {
                "industry_type": "Technology, Internet Services",
                "employee_size": "156,500+",
                "company_website": "https://www.google.com",
                "headquarters": "Mountain View, California, USA",
                "founding_year": "1998",
                "key_products": ["Google Search", "Android", "Google Cloud", "YouTube"],
                "annual_revenue": "$282.84 billion (2023)",
                "key_executives": [
                    {
                        "name": "Sundar Pichai",
                        "role": "CEO",
                        "public_contact": "No direct email provided; see google.com/about for press inquiries"
                    }
                ],
                "official_contacts": {
                    "phone": "650-253-0000",
                    "email": "press@google.com (press inquiries)",
                    "address": "1600 Amphitheatre Parkway, Mountain View, CA"
                },
                "additional_insights": "Parent company is Alphabet Inc."
            },
            # Additional hardcoded entries can be added here.
        }
        return company_info.get(company_name, {})
    
    def _fetch_wikipedia_summary(self, company_name):
        """
        Attempt to fetch a Wikipedia summary for the given company name.
        """
        try:
            summary = wikipedia.summary(company_name, sentences=3)
            return summary
        except Exception as e:
            print(f"Wikipedia lookup failed for {company_name}: {e}")
            return ""
    
    def extract_company_info(self, company_name):
        """
        Extract comprehensive, accurate company information.
        
        If hardcoded data is not available, use the Gemini API with an enhanced prompt that instructs the model to search for 
        the most recent, verified details from reputable online sources.
        
        Required JSON fields:
          - "industry_type"
          - "employee_size"
          - "company_website"
          - "headquarters"
          - "founding_year"
          - "key_products"
          - "annual_revenue"
          - "key_executives" (list of objects with "name", "role", "public_contact")
          - "official_contacts" (object with "phone", "email", "address")
          - "additional_insights"
        
        If a field is not available, return "N/A".
        """
        # Use hardcoded info if available
        hardcoded_info = self._get_hardcoded_company_info(company_name)
        if hardcoded_info:
            return {
                "company_name": company_name,
                **hardcoded_info,
                "source": "Hardcoded",
                "gemini_status": "Not Called"
            }
        
        wiki_summary = self._fetch_wikipedia_summary(company_name)
        
        prompt = f"""
        You are a business data retrieval expert with access to the latest, verified information from reputable sources such as official company websites, trusted financial news outlets, and public databases (including Google's own data also do search in google, meta, bing, linkedin, Ministry of Corporate affairs, google search engine). Your task is to provide comprehensive, accurate, and detailed information about the company "{company_name}".
        
        Guidelines:
        1. Only return factual, verified data.
        2. Search the web for the most recent and accurate details.
        3. Use reputable sources only.
        4. If a detail is not available, return "N/A" (do not guess).
        
        Provide the following fields in your JSON response:
        - "industry_type": The primary industry or sector.
        - "employee_size": The current employee count or range.
        - "company_website": The official website URL.
        - "headquarters": The full headquarters address.
        - "founding_year": The exact year the company was founded.
        - "key_products": An array of 3 to 5 key products or services.
        - "annual_revenue": The most recent annual revenue figure.
        - "key_executives": An array of objects (each with "name", "role", "public_contact").
        - "official_contacts": An object with "phone", "email", and "address".
        - "additional_insights": A short string with any extra relevant remarks.
        
        Also, consider this context from Wikipedia (if available): {wiki_summary}
        
        Return your answer as a valid JSON object with the keys exactly as listed above.
        """
        
        max_retries = 3
        retries = 0
        gemini_status = None
        
        while retries < max_retries:
            try:
                response = self.model.generate_content(prompt)
                gemini_status = "Success"
                company_data = extract_json_from_text(response.text)
                if company_data:
                    source_str = "Gemini API (with Wikipedia context)" if wiki_summary else "Gemini API"
                    company_data["source"] = source_str
                    company_data["gemini_status"] = gemini_status
                    return {"company_name": company_name, **company_data}
                else:
                    gemini_status = "JSON Parsing Failed"
                    print(f"JSON parsing failed for {company_name}. Response: {response.text}")
                    return {
                        "company_name": company_name,
                        "source": "Gemini API",
                        "gemini_status": gemini_status,
                        "error": "Failed to parse JSON from API response."
                    }
            except Exception as e:
                error_message = str(e)
                if "429" in error_message:
                    wait_time = 30
                    print(f"Rate limit exceeded for {company_name}. Waiting {wait_time} seconds (Retry {retries+1}/{max_retries})")
                    time.sleep(wait_time)
                    retries += 1
                else:
                    gemini_status = error_message
                    print(f"Error extracting info for {company_name}: {e}")
                    return {
                        "company_name": company_name,
                        "source": "Gemini API",
                        "gemini_status": gemini_status,
                        "error": error_message
                    }
        
        if wiki_summary:
            return {
                "company_name": company_name,
                "wikipedia_summary": wiki_summary,
                "source": "Wikipedia fallback",
                "gemini_status": gemini_status,
                "error": f"Failed after {max_retries} retries."
            }
        else:
            return {
                "company_name": company_name,
                "source": "Gemini API",
                "gemini_status": gemini_status,
                "error": f"Failed after {max_retries} retries."
            }
    
    def batch_extract(self, companies, progress_callback=None, max_workers=3):
        """
        Process a list of companies concurrently.
        After each company's info is retrieved, call progress_callback(new_result)
        to update the job status in real time.
        """
        results = []
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_company = {
                executor.submit(self.extract_company_info, company): company for company in companies
            }
            for future in as_completed(future_to_company):
                try:
                    result = future.result()
                    results.append(result)
                    if progress_callback:
                        progress_callback(result)
                except Exception as e:
                    company = future_to_company[future]
                    error_result = {
                        "company_name": company,
                        "source": "Unknown",
                        "gemini_status": "Exception",
                        "error": str(e)
                    }
                    results.append(error_result)
                    if progress_callback:
                        progress_callback(error_result)
        return results

# Initialize the extractor with your provided Gemini API key.
EXTRACTOR = CompanyInfoExtractor("Google-Gemini-API-KEY")

def process_job(job_id, companies):
    total = len(companies)
    jobs[job_id]["total"] = total
    jobs[job_id]["start_time"] = time.time()
    jobs[job_id]["results"] = []
    
    def progress_callback(new_result):
        jobs[job_id]["completed"] += 1
        jobs[job_id]["results"].append(new_result)
    
    # Run extraction in the background with progress updates.
    EXTRACTOR.batch_extract(companies, progress_callback=progress_callback, max_workers=3)
    jobs[job_id]["end_time"] = time.time()
    jobs[job_id]["processing_time"] = jobs[job_id]["end_time"] - jobs[job_id]["start_time"]

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/start", methods=["POST"])
def start_job():
    company_input = request.form.get("companies", "")
    companies = [c.strip() for c in re.split(r'[\n,]+', company_input) if c.strip()]
    total = len(companies)
    if total == 0:
        return jsonify({"error": "No companies provided."}), 400
    if total > MAX_COMPANIES:
        return jsonify({"error": f"Too many companies provided. Maximum allowed is {MAX_COMPANIES}."}), 400
    job_id = str(uuid.uuid4())
    jobs[job_id] = {"total": total, "completed": 0}
    thread = threading.Thread(target=process_job, args=(job_id, companies))
    thread.start()
    return jsonify({"job_id": job_id})

@app.route("/progress/<job_id>")
def job_progress(job_id):
    job = jobs.get(job_id)
    if job:
        current_time = time.time()
        elapsed = current_time - job.get("start_time", current_time)
        completed = job.get("completed", 0)
        total = job.get("total", 0)
        estimated_remaining = None
        if completed > 0 and total > 0:
            avg = elapsed / completed
            estimated_remaining = avg * (total - completed)
        response = {
            "total": total,
            "completed": completed,
            "estimated_remaining_time": estimated_remaining,
            "finished": "processing_time" in job,
            "results": job.get("results", [])
        }
        if response["finished"]:
            response["processing_time"] = job.get("processing_time", 0)
        return jsonify(response)
    else:
        return jsonify({"error": "Invalid job id."}), 404

if __name__ == "__main__":
    app.run(debug=True)
