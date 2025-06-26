import os
import requests
from datetime import datetime, timedelta
import feedparser
from newsplease import NewsPlease
from typing import List, Dict, Any, Optional
import json
from pathlib import Path
import time
from urllib.parse import urlparse
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

class NewsFetcher:
    def __init__(self):
        self.newsapi_key = os.getenv('NEWS_API_KEY')
        self.gnews_key = os.getenv('GNEWS_API_KEY')
        self.cache_dir = Path('data/raw_articles')
        self.cache_dir.mkdir(parents=True, exist_ok=True)

    def fetch_newsapi(self, query: str, language: str = 'en', page_size: int = 10) -> List[Dict[str, Any]]:
        """Fetch news from NewsAPI"""
        if not self.newsapi_key:
            print("NewsAPI key not found in environment variables")
            return []
            
        url = 'https://newsapi.org/v2/everything'
        
        # Add quotes to the query to make it more precise
        if ' ' in query and not (query.startswith('"') and query.endswith('"')):
            query = f'"{query}"'
            
        params = {
            'q': query,
            'apiKey': self.newsapi_key,
            'language': language,
            'pageSize': min(max(5, page_size), 100),  # Ensure we get at least 5 results if available, max 100
            'sortBy': 'relevancy',  # Changed from 'publishedAt' to 'relevancy'
            'from': (datetime.now() - timedelta(days=14)).strftime('%Y-%m-%d'),  # Increased to 14 days
            'to': datetime.now().strftime('%Y-%m-%d'),
            'searchIn': 'title,description,content'  # Search in all relevant fields
        }
        
        try:
            print(f"Fetching from NewsAPI with params: { {k: v for k, v in params.items() if k != 'apiKey'} }")
            response = requests.get(url, params=params, timeout=30)
            response.raise_for_status()
            data = response.json()
            
            # Check for API-specific errors
            if data.get('status') == 'error':
                error_msg = data.get('message', 'Unknown error')
                print(f"NewsAPI Error: {error_msg}")
                if 'rate limited' in error_msg.lower():
                    print("You may have exceeded your API rate limit. Please wait before making more requests.")
                return []
                
            articles = data.get('articles', [])
            print(f"Found {len(articles)} articles from NewsAPI for query: {query}")
            
            # Add source information to each article
            for article in articles:
                article['source'] = article.get('source', {}).get('name', 'Unknown')
                article['publishedAt'] = article.get('publishedAt', datetime.now().isoformat())
                article['urlToImage'] = article.get('urlToImage', '')
                
            return articles
            
        except requests.exceptions.RequestException as e:
            print(f"Error fetching from NewsAPI: {e}")
            if hasattr(e, 'response') and e.response is not None:
                print(f"Response status: {e.response.status_code}")
                print(f"Response body: {e.response.text[:500]}")
            return []
        except Exception as e:
            print(f"Unexpected error with NewsAPI: {e}")
            return []

    def fetch_gnews(self, query: str, language: str = 'en', max_results: int = 10) -> List[Dict[str, Any]]:
        """Fetch news from GNews"""
        if not self.gnews_key:
            print("GNews key not found in environment variables")
            return []
            
        # Format query for better results
        if ' ' in query and not query.startswith('"'):
            query = f'"{query}"'
            
        url = 'https://gnews.io/api/v4/search'
        params = {
            'q': query,
            'token': self.gnews_key,
            'lang': language,
            'max': min(max(5, max_results), 100),  # Ensure we get at least 5 results if available, max 100
            'from': (datetime.now() - timedelta(days=14)).strftime('%Y-%m-%dT%H:%M:%SZ'),  # Increased to 14 days
            'to': datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ'),
            'sortby': 'relevance',  # Changed from 'publishedAt' to 'relevance'
            'in': 'title,description,content'  # Search in all relevant fields
        }
        
        try:
            print(f"Fetching from GNews with params: { {k: v for k, v in params.items() if k != 'token'} }")
            response = requests.get(url, params=params, timeout=30)
            response.raise_for_status()
            data = response.json()
            
            # Check for GNews errors
            if 'errors' in data:
                print(f"GNews Error: {data.get('message', 'Unknown error')}")
                if 'quota' in data.get('message', '').lower():
                    print("You may have exceeded your GNews API quota. Please check your account or wait until it resets.")
                return []
            
            # Convert GNews format to match NewsAPI format for consistency
            articles = []
            for article in data.get('articles', []):
                try:
                    # Skip articles without a valid URL
                    if not article.get('url'):
                        continue
                        
                    # Format the article data
                    formatted_article = {
                        'title': article.get('title', '').strip(),
                        'description': article.get('description', '').strip(),
                        'url': article['url'].strip(),
                        'urlToImage': article.get('image', '').strip(),
                        'publishedAt': article.get('publishedAt', datetime.now().isoformat()),
                        'source': {'name': article.get('source', {}).get('name', 'Unknown').strip()},
                        'content': article.get('content', '').strip()
                    }
                    
                    # Only add if we have at least a title and URL
                    if formatted_article['title'] and formatted_article['url']:
                        articles.append(formatted_article)
                        
                except Exception as e:
                    print(f"Error processing GNews article: {e}")
                    continue
            
            print(f"Found {len(articles)} articles from GNews for query: {query}")
            return articles
            
        except requests.exceptions.RequestException as e:
            print(f"Error fetching from GNews: {e}")
            if hasattr(e, 'response') and e.response is not None:
                print(f"Response status: {e.response.status_code}")
                print(f"Response body: {e.response.text[:500]}")
            return []
        except Exception as e:
            print(f"Unexpected error with GNews: {e}")
            return []

    def _parse_article(self, url: str) -> Optional[Dict[str, Any]]:
        """Parse an article from a URL using news-please"""
        try:
            article = NewsPlease.from_url(url, timeout=10)
            if not article:
                return None
                
            return {
                'title': article.title,
                'text': article.maintext,
                'summary': article.description,
                'keywords': article.keywords if hasattr(article, 'keywords') else [],
                'authors': [article.author] if hasattr(article, 'author') and article.author else [],
                'publish_date': article.date_publish.isoformat() if hasattr(article, 'date_publish') and article.date_publish else None,
                'top_image': article.image_url if hasattr(article, 'image_url') else None,
                'images': [article.image_url] if hasattr(article, 'image_url') and article.image_url else [],
                'url': article.url,
                'source_url': article.source_domain,
                'canonical_link': article.url,
                'meta_data': {
                    'description': article.description if hasattr(article, 'description') else '',
                    'language': article.language if hasattr(article, 'language') else 'en',
                    'site_name': article.source_domain
                }
            }
        except Exception as e:
            print(f"Error parsing article from {url}: {e}")
            return None

    def get_article_content(self, url: str) -> str:
        """Extract full article content using news-please"""
        try:
            article = self._parse_article(url)
            return article.get('text', '') if article else ""
        except Exception as e:
            print(f"Error extracting article content from {url}: {e}")
            return ""

    def _get_sector_queries(self, sector: str) -> List[str]:
        """Get specific search queries for specialized sectors"""
        queries = {
            "Semiconductors": [
                "semiconductor industry",
                "chip manufacturing",
                "TSMC Intel AMD Nvidia",
                "semiconductor supply chain",
                "semiconductor technology"
            ],
            "Wearable Technology Sensors": [
                "wearable technology sensors",
                "health monitoring wearables",
                "fitness tracker sensors",
                "wearable medical devices",
                "smart clothing sensors"
            ],
            "Supply Chain": [
                "global supply chain",
                "logistics and supply chain",
                "supply chain management",
                "supply chain technology",
                "supply chain disruption"
            ],
            "Intellectual Property Litigation": [
                "patent litigation",
                "IP lawsuits",
                "intellectual property disputes",
                "trademark infringement",
                "copyright cases"
            ]
        }
        return queries.get(sector, [sector])

    def fetch_news(self, query: str, max_articles: int = 10) -> List[Dict[str, Any]]:
        """Fetch news from multiple sources and combine results"""
        print(f"Fetching news for: {query}")
        
        # Check if we have at least one API key
        if not self.newsapi_key and not self.gnews_key:
            print("No API keys found. Please set NEWS_API_KEY and/or GNEWS_API_KEY environment variables.")
            return []
        
        all_articles = []
        
        # Get specific queries for the sector
        queries = self._get_sector_queries(query)
        
        for q in queries:
            if len(all_articles) >= max_articles:
                break
                
            # Try NewsAPI first if key is available
            if self.newsapi_key:
                print(f"Fetching '{q}' from NewsAPI...")
                news = self.fetch_newsapi(q, page_size=max_articles - len(all_articles))
                all_articles.extend(news)
                time.sleep(1)  # Rate limiting
            
            # If still not enough results, try GNews if key is available
            if self.gnews_key and len(all_articles) < max_articles:
                print(f"Fetching '{q}' from GNews...")
                gnews = self.fetch_gnews(q, max_results=max_articles - len(all_articles))
                all_articles.extend(gnews)
                time.sleep(1)  # Rate limiting
        
        # Deduplicate articles by URL and ensure we don't exceed max_articles
        seen_urls = set()
        unique_articles = []
        for article in all_articles:
            url = article.get('url')
            if url and url not in seen_urls:
                seen_urls.add(url)
                unique_articles.append(article)
                if len(unique_articles) >= max_articles:
                    break
                    
        # Add full content to articles (optional, can be slow)
        for i, article in enumerate(unique_articles):
            try:
                if not article.get('content') or len(str(article.get('content', ''))) < 100:
                    print(f"Extracting content for article {i+1}/{len(unique_articles)}...")
                    full_content = self.get_article_content(article['url'])
                    if full_content:
                        article['content'] = full_content
                    time.sleep(0.5)  # Be respectful to websites
            except Exception as e:
                print(f"Error processing article {i+1}: {e}")
        
        # Cache the results if we have articles
        if unique_articles:
            self._cache_articles(query, unique_articles)
        
        print(f"Found {len(unique_articles)} unique articles for '{query}'")
        return unique_articles
    
    def _cache_articles(self, query: str, articles: List[Dict[str, Any]]) -> None:
        """Cache articles to avoid refetching"""
        try:
            cache_file = self.cache_dir / f"{query.lower().replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.json"
            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump({
                    'query': query,
                    'date': datetime.now().isoformat(),
                    'articles': articles
                }, f, ensure_ascii=False, indent=2)
            print(f"Cached {len(articles)} articles to {cache_file}")
        except Exception as e:
            print(f"Error caching articles: {e}")

    def load_cached_articles(self, query: str) -> List[Dict[str, Any]]:
        """Load articles from cache if available and recent"""
        try:
            cache_file = self.cache_dir / f"{query.lower().replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.json"
            if cache_file.exists():
                with open(cache_file, 'r', encoding='utf-8') as f:
                    cached_data = json.load(f)
                print(f"Loaded {len(cached_data['articles'])} articles from cache")
                return cached_data['articles']
        except Exception as e:
            print(f"Error loading cached articles: {e}")
        return []

    def test_api_keys(self):
        """Test if API keys are working"""
        print("Testing API keys...")
        
        if self.newsapi_key:
            print("Testing NewsAPI...")
            test_articles = self.fetch_newsapi("test", page_size=1)
            if test_articles:
                print("✓ NewsAPI is working")
            else:
                print("✗ NewsAPI failed")
        else:
            print("✗ NewsAPI key not found")
            
        if self.gnews_key:
            print("Testing GNews...")
            test_articles = self.fetch_gnews("test", max_results=1)
            if test_articles:
                print("✓ GNews is working")
            else:
                print("✗ GNews failed")
        else:
            print("✗ GNews key not found")

# Example usage
if __name__ == "__main__":
    
    fetcher = NewsFetcher()
    
    # Test API keys first
    fetcher.test_api_keys()
    
    # Try to load from cache first
    query = "artificial intelligence"
    articles = fetcher.load_cached_articles(query)
    
    # If no cached articles, fetch new ones
    if not articles:
        articles = fetcher.fetch_news(query, max_articles=5)
    
    # Display results
    if articles:
        print(f"\nFound {len(articles)} articles:")
        for i, article in enumerate(articles, 1):
            print(f"\n{i}. {article['title']}")
            print(f"   Source: {article['source']['name']}")
            print(f"   URL: {article['url']}")
            print(f"   Published: {article.get('publishedAt', 'Unknown')}")
            if article.get('description'):
                print(f"   Description: {article['description'][:100]}...")
    else:
        print("No articles found.")