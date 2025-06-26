#!/usr/bin/env python3

print("=== STARTING DEBUG SCRIPT ===")
print("Python is running...")

# Test basic functionality first
print("\n1. Testing basic print functionality...")
for i in range(3):
    print(f"   Test {i+1}: OK")

# Test imports
print("\n2. Testing imports...")
try:
    import os
    print("   ✅ os module imported")
except Exception as e:
    print(f"   ❌ os import error: {e}")

try:
    import sys
    print("   ✅ sys module imported")
    print(f"   Python version: {sys.version}")
except Exception as e:
    print(f"   ❌ sys import error: {e}")

try:
    import requests
    print("   ✅ requests module imported")
except Exception as e:
    print("   ❌ requests not installed. Run: pip install requests")

try:
    from dotenv import load_dotenv
    print("   ✅ python-dotenv imported")
except Exception as e:
    print("   ❌ python-dotenv not installed. Run: pip install python-dotenv")

# Check current directory
print("\n3. Checking current directory...")
try:
    import os
    current_dir = os.getcwd()
    print(f"   Current directory: {current_dir}")
    
    files = os.listdir('.')
    print("   Files in current directory:")
    for file in sorted(files)[:10]:  # Show first 10 files
        print(f"     - {file}")
    
    if len(files) > 10:
        print(f"     ... and {len(files) - 10} more files")
        
except Exception as e:
    print(f"   ❌ Error checking directory: {e}")

# Check for .env file specifically
print("\n4. Checking for .env file...")
try:
    import os
    if os.path.exists('.env'):
        print("   ✅ .env file exists")
        try:
            with open('.env', 'r') as f:
                lines = f.readlines()
            print(f"   .env file has {len(lines)} lines")
            for i, line in enumerate(lines[:5], 1):  # Show first 5 lines
                line = line.strip()
                if line and not line.startswith('#'):
                    # Hide the actual key values for security
                    if '=' in line:
                        key, value = line.split('=', 1)
                        print(f"   Line {i}: {key}={'*' * min(len(value), 10)}")
                    else:
                        print(f"   Line {i}: {line}")
                elif line:
                    print(f"   Line {i}: {line}")
        except Exception as e:
            print(f"   ❌ Error reading .env file: {e}")
    else:
        print("   ❌ .env file does NOT exist")
        print("   You need to create a .env file with your API keys")
except Exception as e:
    print(f"   ❌ Error checking .env file: {e}")

# Test environment variable loading
print("\n5. Testing environment variables...")
try:
    from dotenv import load_dotenv
    load_dotenv()
    print("   ✅ load_dotenv() executed")
    
    import os
    newsapi_key = os.getenv('NEWS_API_KEY')
    gnews_key = os.getenv('GNEWS_API_KEY')
    
    print(f"   NEWS_API_KEY loaded: {newsapi_key is not None}")
    print(f"   GNEWS_API_KEY loaded: {gnews_key is not None}")
    
    if newsapi_key:
        print(f"   NewsAPI key length: {len(newsapi_key)}")
        print(f"   NewsAPI key preview: {newsapi_key[:5]}...{newsapi_key[-3:]}")
    
    if gnews_key:
        print(f"   GNews key length: {len(gnews_key)}")
        print(f"   GNews key preview: {gnews_key[:5]}...{gnews_key[-3:]}")
        
except Exception as e:
    print(f"   ❌ Error with environment variables: {e}")

# Test a simple HTTP request
print("\n6. Testing internet connection...")
try:
    import requests
    response = requests.get('https://httpbin.org/get', timeout=5)
    if response.status_code == 200:
        print("   ✅ Internet connection working")
    else:
        print(f"   ❌ HTTP request failed: {response.status_code}")
except Exception as e:
    print(f"   ❌ Internet connection error: {e}")

# Test NewsAPI if key exists
print("\n7. Testing NewsAPI...")
try:
    from dotenv import load_dotenv
    import os
    import requests
    
    load_dotenv()
    newsapi_key = os.getenv('NEWS_API_KEY')
    
    if newsapi_key:
        print("   Testing NewsAPI with simple request...")
        url = 'https://newsapi.org/v2/top-headlines'
        params = {
            'country': 'us',
            'pageSize': 1,
            'apiKey': newsapi_key
        }
        
        response = requests.get(url, params=params, timeout=10)
        print(f"   Status code: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            total = data.get('totalResults', 0)
            print(f"   ✅ NewsAPI working! Found {total} articles")
        elif response.status_code == 401:
            print("   ❌ NewsAPI key is invalid or unauthorized")
        else:
            print(f"   ❌ NewsAPI error: {response.text[:200]}")
    else:
        print("   ❌ No NewsAPI key found")
        
except Exception as e:
    print(f"   ❌ NewsAPI test error: {e}")

# Test GNews if key exists
print("\n8. Testing GNews...")
try:
    gnews_key = os.getenv('GNEWS_API_KEY')
    
    if gnews_key:
        print("   Testing GNews with simple request...")
        url = 'https://gnews.io/api/v4/top-headlines'
        params = {
            'token': gnews_key,
            'lang': 'en',
            'max': 1
        }
        
        response = requests.get(url, params=params, timeout=10)
        print(f"   Status code: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            articles = data.get('articles', [])
            print(f"   ✅ GNews working! Found {len(articles)} articles")
        elif response.status_code == 401:
            print("   ❌ GNews key is invalid or unauthorized")
        elif response.status_code == 403:
            print("   ❌ GNews key doesn't have required permissions")
        else:
            print(f"   ❌ GNews error: {response.text[:200]}")
    else:
        print("   ❌ No GNews key found")
        
except Exception as e:
    print(f"   ❌ GNews test error: {e}")

print("\n=== DEBUG SCRIPT COMPLETED ===")
print("\nIf you can see this message, the script ran successfully!")
print("If you're still having issues, please share the output above.")