import re
import string
from typing import List, Dict, Any, Tuple
import html
import unicodedata
from bs4 import BeautifulSoup
import spacy
from spacy.lang.en.stop_words import STOP_WORDS
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class TextCleaner:
    """Class for cleaning and preprocessing text data"""
    
    def __init__(self, config: dict):
        """Initialize the text cleaner with configuration"""
        self.config = config
        self.nlp = spacy.load("en_core_web_sm", disable=['parser', 'ner', 'textcat'])
        self.stop_words = set(STOP_WORDS)
        
        # Custom stop words to add
        self.custom_stop_words = {
            'said', 'say', 'says', 'also', 'would', 'could', 
            'may', 'might', 'must', 'shall', 'should', 'via',
            'according', 'like', 'one', 'two', 'three', 'four',
            'five', 'six', 'seven', 'eight', 'nine', 'ten',
            'first', 'second', 'third', 'last', 'new', 'us', 'u'
        }
        
        # Add custom stop words
        for word in self.custom_stop_words:
            self.stop_words.add(word)
            
        # Regular expressions for cleaning
        self.url_regex = re.compile(r'https?\S+|www\.\S+')
        self.email_regex = re.compile(r'\S+@\S+\.\S+')
        self.whitespace_regex = re.compile(r'\s+')
        self.non_ascii_regex = re.compile(r'[^\x00-\x7F]+')
        self.punctuation_table = str.maketrans('', '', string.punctuation)
    
    def clean_text(self, text: str) -> str:
        """
        Clean and normalize text by removing unwanted characters, normalizing whitespace, etc.
        
        Args:
            text: Input text to clean
            
        Returns:
            Cleaned text
        """
        if not text or not isinstance(text, str):
            return ""
            
        try:
            # Convert to string if not already
            text = str(text)
            
            # Decode HTML entities
            text = html.unescape(text)
            
            # Remove HTML tags
            text = BeautifulSoup(text, 'html.parser').get_text(separator=' ')
            
            # Remove URLs
            text = self.url_regex.sub('', text)
            
            # Remove email addresses
            text = self.email_regex.sub('', text)
            
            # Remove non-ASCII characters
            text = self.non_ascii_regex.sub(' ', text)
            
            # Normalize unicode characters
            text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('ascii')
            
            # Convert to lowercase
            text = text.lower()
            
            # Remove punctuation
            text = text.translate(self.punctuation_table)
            
            # Remove numbers
            text = re.sub(r'\d+', '', text)
            
            # Normalize whitespace
            text = self.whitespace_regex.sub(' ', text).strip()
            
            return text
            
        except Exception as e:
            logger.error(f"Error cleaning text: {str(e)}")
            return ""
    
    def lemmatize_text(self, text: str) -> str:
        """
        Lemmatize text using spaCy
        
        Args:
            text: Input text to lemmatize
            
        Returns:
            Lemmatized text
        """
        if not text:
            return ""
            
        try:
            doc = self.nlp(text)
            lemmas = [token.lemma_ for token in doc if not token.is_stop and not token.is_punct]
            return ' '.join(lemmas)
        except Exception as e:
            logger.error(f"Error lemmatizing text: {str(e)}")
            return text
    
    def remove_stopwords(self, text: str) -> str:
        """
        Remove stopwords from text
        
        Args:
            text: Input text
            
        Returns:
            Text with stopwords removed
        """
        if not text:
            return ""
            
        try:
            words = text.split()
            filtered_words = [word for word in words if word.lower() not in self.stop_words]
            return ' '.join(filtered_words)
        except Exception as e:
            logger.error(f"Error removing stopwords: {str(e)}")
            return text
    
    def clean_article(self, article: Dict[str, Any]) -> Dict[str, Any]:
        """
        Clean an article dictionary containing title, content, etc.
        
        Args:
            article: Dictionary containing article data
            
        Returns:
            Dictionary with cleaned fields
        """
        if not article or not isinstance(article, dict):
            return {}
            
        try:
            cleaned = article.copy()
            
            # Clean title
            if 'title' in cleaned:
                cleaned['cleaned_title'] = self.clean_text(cleaned['title'])
                
            # Clean content
            if 'content' in cleaned:
                cleaned['cleaned_content'] = self.clean_text(cleaned['content'])
            
            # Clean description if exists
            if 'description' in cleaned:
                cleaned['cleaned_description'] = self.clean_text(cleaned['description'])
            
            # Add lemmatized versions
            if 'cleaned_content' in cleaned:
                cleaned['lemmatized_content'] = self.lemmatize_text(cleaned['cleaned_content'])
                cleaned['content_no_stopwords'] = self.remove_stopwords(cleaned['cleaned_content'])
            
            return cleaned
            
        except Exception as e:
            logger.error(f"Error cleaning article: {str(e)}")
            return article
    
    def batch_clean_articles(self, articles: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Clean a batch of articles
        
        Args:
            articles: List of article dictionaries
            
        Returns:
            List of cleaned article dictionaries
        """
        if not articles:
            return []
            
        return [self.clean_article(article) for article in articles if article]


def clean_html(html_content: str) -> str:
    """
    Remove HTML tags from content
    
    Args:
        html_content: HTML content to clean
        
    Returns:
        Cleaned text content
    """
    if not html_content:
        return ""
        
    try:
        # Parse HTML and extract text
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Remove script and style elements
        for script in soup(["script", "style"]):
            script.extract()
            
        # Get text
        text = soup.get_text()
        
        # Break into lines and remove leading/trailing whitespace
        lines = (line.strip() for line in text.splitlines())
        # Break multi-headlines into a line each
        chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
        # Drop blank lines
        text = '\n'.join(chunk for chunk in chunks if chunk)
        
        return text
        
    except Exception as e:
        logger.error(f"Error cleaning HTML: {str(e)}")
        return html_content
