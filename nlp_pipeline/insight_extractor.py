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
    
    def extract_key_phrases(self, text: str, top_n: int = 8) -> List[Tuple[str, float]]:
        """
        Extract key phrases from text using noun chunks, named entities, and TF-IDF scoring
        
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
            
            # Get noun chunks and filter out common patterns
            noun_chunks = []
            for chunk in doc.noun_chunks:
                # Filter out single common words unless they're proper nouns
                if len(chunk.text.split()) > 1 or chunk.root.pos_ in ['PROPN', 'NOUN']:
                    # Remove determiners from the beginning
                    while len(chunk) > 1 and chunk[0].dep_ in ['det', 'prep', 'aux']:
                        chunk = chunk[1:]
                    if len(chunk) > 0:
                        noun_chunks.append(chunk.text.lower())
            
            # Get named entities
            entities = [ent.text.lower() for ent in doc.ents if ent.label_ in ['ORG', 'PERSON', 'GPE', 'PRODUCT', 'EVENT']]
            
            # Combine and count frequencies
            all_phrases = noun_chunks + entities
            
            # Calculate TF-IDF like scores
            phrase_scores = {}
            total_phrases = len(all_phrases)
            
            for phrase in set(all_phrases):
                # Term frequency (normalized by document length)
                tf = all_phrases.count(phrase) / total_phrases
                
                # Inverse document frequency (simplified)
                # In a real implementation, you'd use a larger corpus for IDF
                # Here we'll use the length of the phrase as a proxy for specificity
                idf = 1 + (len(phrase.split()) * 0.2)
                
                # Position bonus (earlier mentions are often more important)
                first_occurrence = text.lower().find(phrase)
                position_bonus = 1.0
                if first_occurrence >= 0:
                    position = first_occurrence / len(text)
                    position_bonus = 1.5 - position  # Higher for earlier mentions
                
                # Combine scores
                score = tf * idf * position_bonus
                
                # Bonus for phrases that contain important terms
                important_terms = {'announce', 'launch', 'develop', 'new', 'report', 'study', 'find', 'show', 'reveal'}
                if any(term in phrase for term in important_terms):
                    score *= 1.5
                
                phrase_scores[phrase] = score
            
            # Sort phrases by score and take top N
            top_phrases = sorted(phrase_scores.items(), key=lambda x: x[1], reverse=True)[:top_n]
            
            # Normalize scores to 0-1 range
            max_score = max([score for _, score in top_phrases], default=1.0)
            if max_score > 0:
                scored_phrases = [(phrase, score/max_score) for phrase, score in top_phrases]
            else:
                scored_phrases = [(phrase, score) for phrase, score in top_phrases]
            
            return scored_phrases
            
        except Exception as e:
            logger.error(f"Error extracting key phrases: {str(e)}")
            return []
    
    def _generate_detailed_summary(self, text: str, max_sentences: int = 8) -> str:
        """
        Generate a detailed summary of the text using extractive summarization
        
        Args:
            text: Input text to summarize
            max_sentences: Maximum number of sentences for the summary
            
        Returns:
            Generated summary as a string
        """
        if not text:
            return ""
            
        try:
            # Process the text with spaCy
            doc = self.nlp(text)
            
            # Extract sentences and their importance scores
            sentence_scores = []
            for sent in doc.sents:
                # Calculate score based on sentence length and position
                # Give more weight to sentences that contain named entities or key terms
                entity_score = len([ent for ent in sent.ents]) * 0.5
                term_score = sum(1 for token in sent if token.is_alpha and token.lemma_.lower() in {
                    'announce', 'launch', 'develop', 'introduce', 'partner', 'collaborate',
                    'increase', 'grow', 'expand', 'invest', 'fund', 'raise', 'acquire',
                    'merge', 'release', 'report', 'according', 'show', 'reveal', 'indicate'
                })
                
                # Position-based scoring (first and last sentences are often important)
                position = sent.start / len(list(doc.sents))
                position_score = 1.5 - abs(position - 0.5)  # Higher for start/end
                
                # Combine scores
                score = (entity_score + term_score + position_score) * len(sent.text.split())
                sentence_scores.append((sent.text.strip(), score, sent.start))
            
            # Sort sentences by score (highest first)
            sentence_scores.sort(key=lambda x: x[1], reverse=True)
            
            # Take top sentences and reorder them by original position
            top_sentences = sorted(sentence_scores[:max_sentences], key=lambda x: x[2])
            
            # Join sentences to form the summary
            summary = ' '.join([s[0] for s in top_sentences])
            
            # Ensure the summary is not too short
            if len(summary.split()) < 50 and len(text.split()) > 100:  # If summary is too short but text is long
                # Fall back to first few sentences
                sentences = [sent.text.strip() for sent in doc.sents]
                summary = ' '.join(sentences[:min(5, len(sentences))])
                
            return summary
            
        except Exception as e:
            logger.error(f"Error generating summary: {str(e)}")
            # Fallback to first 2 sentences
            doc = self.nlp(text)
            sentences = [sent.text.strip() for sent in doc.sents]
            return ' '.join(sentences[:2]) if len(sentences) > 1 else text[:500] + '...'
    
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
            
            # Generate detailed summary
            summary = self._generate_detailed_summary(text)
            
            # Extract entities
            entities = self.extract_entities(text)
            
            # Extract insight categories
            categories = self.extract_insight_categories(text)
            
            # Extract key phrases
            key_phrases = self.extract_key_phrases(text, top_n=8)  # Get more key phrases for better context
            
            # Prepare result
            insights = {
                'entities': entities,
                'categories': categories,
                'key_phrases': key_phrases,
                'primary_category': categories[0] if categories else 'other',
                'relevance_score': self._calculate_relevance_score(article, categories, key_phrases),
                'summary': summary,  # Add the generated summary
                'detailed_analysis': {
                    'key_points': [phrase[0] for phrase in key_phrases[:5]],
                    'mentioned_organizations': entities.get('ORG', [])[:5],
                    'mentioned_people': entities.get('PERSON', [])[:3],
                    'mentioned_locations': entities.get('GPE', [])[:3]
                }
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
