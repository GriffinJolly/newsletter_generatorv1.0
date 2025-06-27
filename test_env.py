import os
from dotenv import load_dotenv

def test_env_vars():
    # Load environment variables from .env file
    load_dotenv()
    
    # Get the API keys
    news_api_key = os.getenv('NEWS_API_KEY')
    gnews_api_key = os.getenv('GNEWS_API_KEY')
    
    # Print the results
    print("Environment Variables Test")
    print("=" * 30)
    print(f"NEWS_API_KEY: {'✅ Found' if news_api_key else '❌ Not found'}")
    print(f"GNEWS_API_KEY: {'✅ Found' if gnews_api_key else '❌ Not found'}")
    
    if news_api_key and news_api_key.startswith('your_'):
        print("\n⚠️  WARNING: NEWS_API_KEY appears to be using the default value")
    if gnews_api_key and gnews_api_key.startswith('your_'):
        print("⚠️  WARNING: GNEWS_API_KEY appears to be using the default value")

if __name__ == "__main__":
    test_env_vars()
