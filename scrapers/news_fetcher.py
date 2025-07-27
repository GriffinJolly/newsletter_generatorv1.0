import os
import json
import requests
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional
from pathlib import Path
import feedparser
import time
from newspaper import Article
from bs4 import BeautifulSoup
from urllib.parse import urlparse
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class NewsFetcher:
    """Base class for news fetchers"""
    
    def __init__(self, config: dict):
        self.config = config
        self.data_dir = Path(config.get('data_dir', 'data/raw_articles'))
        self.data_dir.mkdir(parents=True, exist_ok=True)
    
    def fetch(self, query: str, max_results: int = 50) -> List[Dict[str, Any]]:
        """Fetch news articles based on query"""
        raise NotImplementedError("Subclasses must implement fetch method")
    
    def save_articles(self, articles: List[Dict[str, Any]], source: str) -> str:
        """Save articles to disk with metadata"""
        if not articles:
            return ""
            
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = self.data_dir / f"{source}_{timestamp}.json"
        
        # Add metadata
        for article in articles:
            article['_metadata'] = {
                'source': source,
                'fetched_at': datetime.now().isoformat(),
                'version': '1.0'
            }
        
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(articles, f, indent=2, ensure_ascii=False)
        
        return str(filename)


class NewsAPIFetcher(NewsFetcher):
    """Fetches news from NewsAPI"""
    
    def __init__(self, config: dict):
        super().__init__(config)
        self.api_key = os.getenv('NEWS_API_KEY')
        self.base_url = "https://newsapi.org/v2"
    
    def fetch(self, query: str, max_results: int = 50) -> List[Dict[str, Any]]:
        if not self.api_key:
            logger.error("NewsAPI key not found in environment variables")
            return []
        
        params = {
            'q': query,
            'apiKey': self.api_key,
            'pageSize': min(max_results, 100),  # Max 100 results per page
            'language': 'en',
            'sortBy': 'publishedAt'
        }
        
        try:
            response = requests.get(f"{self.base_url}/everything", params=params)
            response.raise_for_status()
            data = response.json()
            
            articles = []
            for article in data.get('articles', [])[:max_results]:
                processed = {
                    'title': article.get('title', ''),
                    'url': article.get('url', ''),
                    'source': article.get('source', {}).get('name', 'Unknown'),
                    'published_at': article.get('publishedAt', ''),
                    'content': article.get('content', ''),
                    'description': article.get('description', ''),
                    'author': article.get('author', ''),
                    'url_to_image': article.get('urlToImage', '')
                }
                # --- Full-text extraction logic ---
                try:
                    content = processed['content'] or ''
                    url = processed['url']
                    if url and (not content or len(content) < 500):
                        from newspaper import Article as NPArticle
                        np_article = NPArticle(url)
                        np_article.download()
                        np_article.parse()
                        full_text = np_article.text.strip()
                        if full_text and len(full_text) > len(content):
                            processed['content'] = full_text
                            logger.info(f"Full-text extracted for: {url} (length={len(full_text)})")
                        else:
                            logger.info(f"No longer full-text found for: {url}")
                except Exception as ex:
                    logger.warning(f"Full-text extraction failed for {processed.get('url','')}: {ex}")
                # --- End full-text extraction ---
                articles.append(processed)
            
            self.save_articles(articles, 'newsapi')
            return articles
            
        except Exception as e:
            logger.error(f"Error fetching from NewsAPI: {str(e)}")
            return []


class GNewsFetcher(NewsFetcher):
    """Fetches news from GNews API"""
    
    def __init__(self, config: dict):
        super().__init__(config)
        self.api_key = os.getenv('GNEWS_API_KEY')
        self.base_url = "https://gnews.io/api/v4"
    
    def fetch(self, query: str, max_results: int = 50) -> List[Dict[str, Any]]:
        if not self.api_key:
            logger.error("GNews API key not found in environment variables")
            return []
        
        params = {
            'q': query,
            'token': self.api_key,
            'lang': 'en',
            'country': 'us',
            'max': min(max_results, 100),  # Max 100 results
            'in': 'title,description,content'
        }
        
        try:
            response = requests.get(f"{self.base_url}/search", params=params)
            response.raise_for_status()
            data = response.json()
            
            articles = []
            for article in data.get('articles', [])[:max_results]:
                processed = {
                    'title': article.get('title', ''),
                    'url': article.get('url', ''),
                    'source': article.get('source', {}).get('name', 'Unknown'),
                    'published_at': article.get('publishedAt', ''),
                    'content': article.get('content', ''),
                    'description': article.get('description', ''),
                    'image': article.get('image', ''),
                    'source_url': article.get('url', '')
                }
                # --- Full-text extraction logic ---
                try:
                    content = processed['content'] or ''
                    url = processed['url']
                    if url and (not content or len(content) < 500):
                        from newspaper import Article as NPArticle
                        np_article = NPArticle(url)
                        np_article.download()
                        np_article.parse()
                        full_text = np_article.text.strip()
                        if full_text and len(full_text) > len(content):
                            processed['content'] = full_text
                            logger.info(f"Full-text extracted for: {url} (length={len(full_text)})")
                        else:
                            logger.info(f"No longer full-text found for: {url}")
                except Exception as ex:
                    logger.warning(f"Full-text extraction failed for {processed.get('url','')}: {ex}")
                # --- End full-text extraction ---
                articles.append(processed)
            
            self.save_articles(articles, 'gnews')
            return articles
            
        except Exception as e:
            logger.error(f"Error fetching from GNews: {str(e)}")
            return []


class SECFetcher(NewsFetcher):
    """Fetches SEC filings from EDGAR RSS feeds"""
    
    def __init__(self, config: dict):
        super().__init__(config)
        self.base_url = "https://www.sec.gov/Archives/edgar/xbrlrss"
    
    def fetch(self, query: str = "", max_results: int = 20) -> List[Dict[str, Any]]:
        try:
            # Get the main RSS feed
            feed_url = f"{self.base_url}/all.xml"
            feed = feedparser.parse(feed_url)
            
            articles = []
            for entry in feed.entries[:max_results]:
                # Extract company name from title (format: "COMPANY NAME: Filing Type")
                title_parts = entry.get('title', '').split(':')
                company = title_parts[0].strip() if len(title_parts) > 1 else 'Unknown Company'
                filing_type = title_parts[1].strip() if len(title_parts) > 1 else 'Filing'
                
                # Get filing URL (replace -index.htm with .txt for the actual filing)
                filing_url = entry.get('link', '').replace('-index.htm', '.txt')
                
                processed = {
                    'title': f"{company} - {filing_type}",
                    'url': filing_url,
                    'source': 'SEC EDGAR',
                    'published_at': entry.get('published', ''),
                    'content': entry.get('summary', ''),
                    'description': f"{filing_type} filing from {company}",
                    'company': company,
                    'filing_type': filing_type
                }
                articles.append(processed)
            
            self.save_articles(articles, 'sec_edgar')
            return articles
            
        except Exception as e:
            logger.error(f"Error fetching from SEC EDGAR: {str(e)}")
            return []


class WebScraper(NewsFetcher):
    """Scrapes news articles from given URLs"""
    
    def __init__(self, config: dict):
        super().__init__(config)
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
    
    def fetch_from_urls(self, urls: List[str]) -> List[Dict[str, Any]]:
        """Fetch and process articles from a list of URLs"""
        articles = []
        
        for url in urls:
            try:
                # Use newspaper3k to extract article content
                article = Article(url, headers=self.headers)
                article.download()
                article.parse()
                
                processed = {
                    'title': article.title,
                    'url': url,
                    'source': article.source_url,
                    'published_at': article.publish_date.isoformat() if article.publish_date else '',
                    'content': article.text,
                    'description': article.meta_description,
                    'authors': article.authors,
                    'top_image': article.top_image,
                    'keywords': article.keywords,
                    'summary': article.summary
                }
                articles.append(processed)
                
                # Be nice to servers
                time.sleep(1)
                
            except Exception as e:
                logger.error(f"Error scraping {url}: {str(e)}")
                continue
        
        if articles:
            self.save_articles(articles, 'web_scraper')
            
        return articles
    
    def search_and_fetch(self, query: str, max_results: int = 10) -> List[Dict[str, Any]]:
        """Search the web for articles matching the query and fetch them"""
        # This is a simplified example - in practice, you'd use a search API
        # or scrape search results from a search engine
        logger.warning("Web search not implemented. Please provide direct URLs.")
        return []


def get_fetchers(config: dict) -> List[NewsFetcher]:
    """Initialize and return all available fetchers"""
    fetchers = []
    
    if os.getenv('NEWS_API_KEY'):
        fetchers.append(NewsAPIFetcher(config))
    
    if os.getenv('GNEWS_API_KEY'):
        fetchers.append(GNewsFetcher(config))
    
    # Always include SEC fetcher as it doesn't require an API key
    fetchers.append(SECFetcher(config))
    fetchers.append(WebScraper(config))
    
    return fetchers
