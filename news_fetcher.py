import os
import requests
from datetime import datetime, timedelta
import feedparser
from newspaper import Article
from typing import List, Dict, Any
import json
from pathlib import Path

class NewsFetcher:
    def __init__(self):
        self.newsapi_key = os.getenv('NEWS_API_KEY')
        self.gnews_key = os.getenv('GNEWS_API_KEY')
        self.cache_dir = Path('data/raw_articles')
        self.cache_dir.mkdir(parents=True, exist_ok=True)

    def fetch_newsapi(self, query: str, language: str = 'en', page_size: int = 10) -> List[Dict[str, Any]]:
        """Fetch news from NewsAPI"""
        url = 'https://newsapi.org/v2/everything'
        params = {
            'q': query,
            'apiKey': self.newsapi_key,
            'language': language,
            'pageSize': page_size,
            'sortBy': 'publishedAt',
            'from_param': (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
        }
        
        try:
            response = requests.get(url, params=params)
            response.raise_for_status()
            return response.json().get('articles', [])
        except Exception as e:
            print(f"Error fetching from NewsAPI: {e}")
            return []

    def fetch_gnews(self, query: str, language: str = 'en', max_results: int = 10) -> List[Dict[str, Any]]:
        """Fetch news from GNews"""
        url = 'https://gnews.io/api/v4/search'
        params = {
            'q': query,
            'token': self.gnews_key,
            'lang': language,
            'max': max_results,
            'from': (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%dT%H:%M:%SZ'),
            'to': datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ'),
            'sortby': 'publishedAt'
        }
        
        try:
            response = requests.get(url, params=params)
            response.raise_for_status()
            articles = response.json().get('articles', [])
            
            # Process articles to match our standard format
            processed = []
            for article in articles:
                processed.append({
                    'title': article.get('title', ''),
                    'description': article.get('description', ''),
                    'url': article.get('url', ''),
                    'publishedAt': article.get('publishedAt', ''),
                    'source': {'name': article.get('source', {}).get('name', 'Unknown')},
                    'content': article.get('content', '')
                })
            return processed
        except Exception as e:
            print(f"Error fetching from GNews: {e}")
            return []

    def get_article_content(self, url: str) -> str:
        """Extract full article content using newspaper3k"""
        try:
            article = Article(url)
            article.download()
            article.parse()
            return article.text
        except Exception as e:
            print(f"Error extracting article content: {e}")
            return ""

    def fetch_news(self, query: str, max_articles: int = 10) -> List[Dict[str, Any]]:
        """Fetch news from multiple sources and combine results"""
        print(f"Fetching news for: {query}")
        
        # Get articles from both sources
        newsapi_articles = self.fetch_newsapi(query, page_size=max_articles//2)
        gnews_articles = self.fetch_gnews(query, max_results=max_articles//2)
        
        # Combine and deduplicate articles by URL
        all_articles = {}
        for article in newsapi_articles + gnews_articles:
            url = article.get('url')
            if url and url not in all_articles:
                all_articles[url] = article
        
        # Convert back to list and limit to max_articles
        articles = list(all_articles.values())[:max_articles]
        
        # Add full content to articles
        for article in articles:
            if not article.get('content'):
                article['content'] = self.get_article_content(article['url'])
        
        # Cache the results
        self._cache_articles(query, articles)
        
        return articles
    
    def _cache_articles(self, query: str, articles: List[Dict[str, Any]]) -> None:
        """Cache articles to avoid refetching"""
        cache_file = self.cache_dir / f"{query.lower().replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.json"
        with open(cache_file, 'w', encoding='utf-8') as f:
            json.dump({
                'query': query,
                'date': datetime.now().isoformat(),
                'articles': articles
            }, f, ensure_ascii=False, indent=2)

# Example usage
if __name__ == "__main__":
    import os
    from dotenv import load_dotenv
    load_dotenv()
    
    fetcher = NewsFetcher()
    articles = fetcher.fetch_news("artificial intelligence", max_articles=5)
    for i, article in enumerate(articles, 1):
        print(f"{i}. {article['title']} - {article['source']['name']}")
        print(f"   {article['url']}\n")
