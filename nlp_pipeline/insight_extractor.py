import re
import logging
from typing import List, Dict, Any, Tuple, Set, Optional
import spacy
from spacy.matcher import PhraseMatcher
from spacy.tokens import Span, Doc
import numpy as np
from collections import defaultdict, Counter
import json
from pathlib import Path

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class InsightExtractor:
    """Extract key insights from text using NLP techniques"""
    
    def __init__(self, config: dict):
        """Initialize the insight extractor with configuration"""
        self.config = config
        
        # Load NLP model
        try:
            self.nlp = spacy.load("en_core_web_lg")
        except OSError:
            logger.error("SpaCy model 'en_core_web_lg' not found. Please install it with 'python -m spacy download en_core_web_lg'")
            raise
            
        # Initialize phrase matcher for key terms
        self.matcher = PhraseMatcher(self.nlp.vocab, attr="LOWER")
        
        # Load patterns for different insight types
        self._load_insight_patterns()
        
        # Define insight categories
        self.insight_categories = {
            'merger': ['merger', 'acquisition', 'acquire', 'takeover', 'buyout'],
            'partnership': ['partnership', 'collaboration', 'alliance', 'joint venture', 'team up'],
            'funding': ['funding', 'investment', 'raise', 'series a', 'series b', 'series c', 'funding round'],
            'product': ['launch', 'release', 'new product', 'announce', 'introduce'],
            'leadership': ['appoint', 'hire', 'join', 'name', 'promote', 'resign', 'step down'],
            'financial': ['revenue', 'profit', 'loss', 'earnings', 'financial results', 'quarterly results'],
            'regulation': ['regulation', 'compliance', 'lawsuit', 'settlement', 'fine', 'investigation']
        }
        
        # Add patterns to matcher
        for category, terms in self.insight_categories.items():
            patterns = [self.nlp.make_doc(term) for term in terms]
            self.matcher.add(category.upper(), patterns)
    
    def _load_insight_patterns(self, patterns_file: Optional[str] = None) -> None:
        """
        Load patterns for insight extraction from file or use defaults
        
        Args:
            patterns_file: Path to JSON file containing patterns
        """
        if patterns_file and Path(patterns_file).exists():
            try:
                with open(patterns_file, 'r') as f:
                    self.insight_patterns = json.load(f)
                logger.info(f"Loaded insight patterns from {patterns_file}")
                return
            except Exception as e:
                logger.error(f"Error loading patterns file: {str(e)}")
        
        # Default patterns if file not found or error loading
        self.insight_patterns = {
            'merger': [
                {'label': 'MERGERS_ACQUISITIONS', 'pattern': [{'LOWER': 'acquire'}, {'POS': 'DET', 'OP': '?'}, {}]},
                {'label': 'MERGERS_ACQUISITIONS', 'pattern': [{'LEMMA': 'merge', 'POS': 'VERB'}]},
                {'label': 'MERGERS_ACQUISITIONS', 'pattern': [{'LOWER': 'takeover'}]},
            ],
            'partnership': [
                {'label': 'PARTNERSHIPS', 'pattern': [{'LEMMA': 'partner', 'POS': 'VERB'}]},
                {'label': 'PARTNERSHIPS', 'pattern': [{'LOWER': 'joint'}, {'LOWER': 'venture'}]},
                {'label': 'PARTNERSHIPS', 'pattern': [{'LOWER': 'collaborat'}, {'POS': 'ADP', 'OP': '?'}, {}]},
            ],
            'funding': [
                {'label': 'FUNDING', 'pattern': [{'LOWER': 'raise'}, {'POS': 'DET', 'OP': '?'}, {'ENT_TYPE': 'MONEY'}]},
                {'label': 'FUNDING', 'pattern': [{'LOWER': 'series'}, {'LOWER': {'IN': ['a', 'b', 'c', 'd']}}]},
                {'label': 'FUNDING', 'pattern': [{'LOWER': 'funding'}, {'LOWER': 'round'}]},
            ]
        }
    
    def extract_entities(self, text: str) -> Dict[str, List[str]]:
        """
        Extract named entities from text
        
        Args:
            text: Input text
            
        Returns:
            Dictionary of entity types and their values
        """
        if not text:
            return {}
            
        try:
            doc = self.nlp(text)
            entities = defaultdict(list)
            
            for ent in doc.ents:
                if ent.label_ in ['ORG', 'PERSON', 'GPE', 'NORP', 'FAC', 'PRODUCT', 'EVENT']:
                    entities[ent.label_].append(ent.text)
            
            # Remove duplicates while preserving order
            for key in entities:
                seen = set()
                entities[key] = [x for x in entities[key] if not (x in seen or seen.add(x))]
            
            return dict(entities)
            
        except Exception as e:
            logger.error(f"Error extracting entities: {str(e)}")
            return {}
    
    def extract_insight_categories(self, text: str) -> List[str]:
        """
        Extract insight categories from text
        
        Args:
            text: Input text
            
        Returns:
            List of insight categories
        """
        if not text:
            return []
            
        try:
            doc = self.nlp(text.lower())
            matches = self.matcher(doc)
            
            categories = set()
            for match_id, start, end in matches:
                category = self.nlp.vocab.strings[match_id].lower()
                categories.add(category)
            
            return list(categories)
            
        except Exception as e:
            logger.error(f"Error extracting insight categories: {str(e)}")
            return []
    
    def extract_key_phrases(self, text: str, top_n: int = 5) -> List[Tuple[str, float]]:
        """
        Extract key phrases from text using noun chunks and named entities
        
        Args:
            text: Input text
            top_n: Number of key phrases to return
            
        Returns:
            List of (phrase, score) tuples
        """
        if not text:
            return []
            
        try:
            doc = self.nlp(text)
            
            # Count noun chunks
            noun_chunks = list(doc.noun_chunks)
            noun_chunk_texts = [chunk.text.lower() for chunk in noun_chunks]
            
            # Count named entities
            entities = [ent.text.lower() for ent in doc.ents]
            
            # Combine and count frequencies
            all_phrases = noun_chunk_texts + entities
            phrase_counts = Counter(all_phrases)
            
            # Get top phrases by frequency
            top_phrases = phrase_counts.most_common(top_n)
            
            # Normalize scores to 0-1 range
            max_count = max([count for _, count in top_phrases], default=1)
            scored_phrases = [(phrase, count/max_count) for phrase, count in top_phrases]
            
            return scored_phrases
            
        except Exception as e:
            logger.error(f"Error extracting key phrases: {str(e)}")
            return []
    
    def extract_insights(self, article: Dict[str, Any]) -> Dict[str, Any]:
        """
        Extract insights from an article
        
        Args:
            article: Dictionary containing article data
            
        Returns:
            Dictionary with extracted insights
        """
        if not article or not isinstance(article, dict):
            return {}
            
        try:
            # Get text to analyze (prefer cleaned content if available)
            text = article.get('cleaned_content', article.get('content', ''))
            if not text:
                return {}
            
            # Extract entities
            entities = self.extract_entities(text)
            
            # Extract insight categories
            categories = self.extract_insight_categories(text)
            
            # Extract key phrases
            key_phrases = self.extract_key_phrases(text)
            
            # Prepare result
            insights = {
                'entities': entities,
                'categories': categories,
                'key_phrases': key_phrases,
                'primary_category': categories[0] if categories else 'other',
                'relevance_score': self._calculate_relevance_score(article, categories, key_phrases)
            }
            
            # Add to article
            article['insights'] = insights
            
            return insights
            
        except Exception as e:
            logger.error(f"Error extracting insights: {str(e)}")
            return {}
    
    def _calculate_relevance_score(self, article: Dict[str, Any], 
                                 categories: List[str], 
                                 key_phrases: List[Tuple[str, float]]) -> float:
        """
        Calculate a relevance score for the article based on insights
        
        Args:
            article: Article data
            categories: List of insight categories
            key_phrases: List of (phrase, score) tuples
            
        Returns:
            Relevance score (0-1)
        """
        score = 0.0
        
        # Base score based on categories
        if categories:
            score += 0.3  # Base score for having any category
            
        # Add score for key phrases
        if key_phrases:
            # Take average of top 3 key phrase scores
            top_scores = [s for _, s in key_phrases[:3]]
            score += 0.2 * (sum(top_scores) / len(top_scores) if top_scores else 0)
        
        # Boost for certain entity types
        entities = article.get('insights', {}).get('entities', {})
        if 'ORG' in entities and len(entities['ORG']) > 0:
            score += 0.2
        if 'PERSON' in entities and len(entities['PERSON']) > 0:
            score += 0.1
        if 'GPE' in entities and len(entities['GPE']) > 0:
            score += 0.1
        
        # Cap score at 1.0
        return min(1.0, score)
    
    def process_articles(self, articles: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Process a list of articles to extract insights
        
        Args:
            articles: List of article dictionaries
            
        Returns:
            List of processed articles with insights
        """
        if not articles:
            return []
            
        processed = []
        for article in articles:
            processed_article = article.copy()
            self.extract_insights(processed_article)
            processed.append(processed_article)
            
        return processed
