import os
import sys
import requests
import json
from datetime import datetime, timedelta
from pathlib import Path
from dotenv import load_dotenv

def print_header(title):
    print("\n" + "="*50)
    print(f"{title:^50}")
    print("="*50)

def check_env_file():
    print_header("1. CHECKING ENVIRONMENT")
    env_path = Path('.env')
    if not env_path.exists():
        print("❌ .env file not found in current directory")
        print(f"Current directory: {Path.cwd()}")
        return False
    
    print("✅ .env file found")
    
    # Load environment variables
    load_dotenv()
    
    # Check for required API keys
    newsapi_key = os.getenv('NEWS_API_KEY')
    gnews_key = os.getenv('GNEWS_API_KEY')
    
    print(f"NEWS_API_KEY: {'✅ Found' if newsapi_key else '❌ Missing'}")
    print(f"GNEWS_API_KEY: {'✅ Found' if gnews_key else '❌ Missing'})")
    
    return bool(newsapi_key or gnews_key)

def test_internet_connection():
    print_header("2. TESTING INTERNET CONNECTION")
    try:
        response = requests.get('https://www.google.com', timeout=10)
        print("✅ Internet connection is working")
        return True
    except Exception as e:
        print(f"❌ No internet connection: {e}")
        return False

def test_newsapi():
    print_header("3. TESTING NEWSAPI")
    api_key = os.getenv('NEWS_API_KEY')
    
    if not api_key:
        print("❌ NEWS_API_KEY not found in environment variables")
        return False
    
    print(f"Testing with NEWS_API_KEY: {api_key[:5]}...{api_key[-3:]}")
    
    url = 'https://newsapi.org/v2/everything'
    params = {
        'q': 'technology',
        'apiKey': api_key,
        'pageSize': 1,
        'sortBy': 'publishedAt'
    }
    
    try:
        print("Sending request to NewsAPI...")
        response = requests.get(url, params=params, timeout=15)
        print(f"Status Code: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            print("✅ NewsAPI is working!")
            print(f"Found {data.get('totalResults', 0)} total articles")
            if 'articles' in data and data['articles']:
                article = data['articles'][0]
                print("\nSample Article:")
                print(f"Title: {article.get('title')}")
                print(f"Source: {article.get('source', {}).get('name', 'Unknown')}")
                print(f"URL: {article.get('url')}")
            return True
        elif response.status_code == 401:
            print("❌ NewsAPI key is invalid or unauthorized")
            print("Response:", response.text[:200])
        elif response.status_code == 429:
            print("❌ Rate limit exceeded for NewsAPI")
        else:
            print(f"❌ Error from NewsAPI: {response.status_code}")
            print("Response:", response.text[:200])
    except Exception as e:
        print(f"❌ Error testing NewsAPI: {e}")
    
    return False

def test_gnews():
    print_header("4. TESTING GNEWS")
    api_key = os.getenv('GNEWS_API_KEY')
    
    if not api_key:
        print("❌ GNEWS_API_KEY not found in environment variables")
        return False
    
    print(f"Testing with GNEWS_API_KEY: {api_key[:5]}...{api_key[-3:]}")
    
    url = 'https://gnews.io/api/v4/top-headlines'
    params = {
        'token': api_key,
        'lang': 'en',
        'max': 1,
        'category': 'technology'
    }
    
    try:
        print("Sending request to GNews...")
        response = requests.get(url, params=params, timeout=15)
        print(f"Status Code: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            print("✅ GNews is working!")
            if 'articles' in data and data['articles']:
                article = data['articles'][0]
                print("\nSample Article:")
                print(f"Title: {article.get('title')}")
                print(f"Source: {article.get('source', {}).get('name', 'Unknown')}")
                print(f"URL: {article.get('url')}")
            return True
        elif response.status_code == 401:
            print("❌ GNews key is invalid or unauthorized")
            print("Response:", response.text[:200])
        elif response.status_code == 429:
            print("❌ Rate limit exceeded for GNews")
        else:
            print(f"❌ Error from GNews: {response.status_code}")
            print("Response:", response.text[:200])
    except Exception as e:
        print(f"❌ Error testing GNews: {e}")
    
    return False

def main():
    print_header("NEWS API DEBUGGER")
    print(f"Running at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Python version: {sys.version.split()[0]}")
    print(f"Current directory: {Path.cwd()}")
    
    # Check environment and API keys
    env_ok = check_env_file()
    if not env_ok:
        print("\n❌ Please check your .env file and API keys")
        return
    
    # Test internet connection
    if not test_internet_connection():
        print("\n❌ Please check your internet connection")
        return
    
    # Test APIs
    newsapi_ok = test_newsapi()
    gnews_ok = test_gnews()
    
    print_header("DEBUG SUMMARY")
    if newsapi_ok or gnews_ok:
        print("✅ At least one API is working!")
    else:
        print("❌ None of the APIs are working. Please check the errors above.")
    
    print("\nTroubleshooting Tips:")
    print("1. Verify your API keys in the .env file")
    print("2. Check your internet connection")
    print("3. Visit the API provider's website to check your API key status")
    print("4. Make sure your API key has the required permissions")
    print("5. Check if you've reached your API rate limits")

if __name__ == "__main__":
    main()
