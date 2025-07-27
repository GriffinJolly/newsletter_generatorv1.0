import sys
# Patch for Streamlit watcher bug with torch._classes
sys.modules['torch._classes'] = None
import streamlit as st
import os
from datetime import datetime, timedelta
import yaml
from pathlib import Path
from pptx.util import Inches, Pt
import sys
import logging
from typing import Dict, List, Any, Optional
import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
from wordcloud import WordCloud
import base64
from io import BytesIO

# Add parent directory to path
sys.path.append(str(Path(__file__).parent.parent))
from news_fetcher import NewsFetcher

# Initialize the news fetcher
fetcher = NewsFetcher()

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def load_config():
    """Load configuration from config.yaml"""
    config_path = Path(__file__).parent.parent / 'config.yaml'
    with open(config_path, 'r') as f:
        return yaml.safe_load(f)

CONFIG = load_config()

# Initialize session state
if 'generated' not in st.session_state:
    st.session_state.generated = False
if 'newsletter_path' not in st.session_state:
    st.session_state.newsletter_path = ""
if 'news_data' not in st.session_state:
    st.session_state.news_data = {}

# Set page config
st.set_page_config(
    page_title="Sector Newsletter Generator",
    page_icon="📰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main {
        max-width: 1200px;
        padding: 2rem;
    }
    .stButton>button {
        width: 100%;
        padding: 0.5rem;
        border-radius: 5px;
        background-color: #4CAF50;
        color: white;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #45a049;
    }
    .stTextInput>div>div>input {
        padding: 0.5rem;
        border-radius: 5px;
        border: 1px solid #ccc;
    }
    .stSelectbox>div>div>div {
        padding: 0.5rem;
        border-radius: 5px;
        border: 1px solid #ccc;
    }
    .news-card {
        border-left: 4px solid #4CAF50;
        padding: 1rem;
        margin: 0.5rem 0;
        background-color: #f9f9f9;
        border-radius: 0 5px 5px 0;
    }
    .news-card h4 {
        margin-top: 0;
        color: #2c3e50;
    }
    .news-card p {
        margin-bottom: 0.5rem;
        color: #34495e;
    }
    .news-source {
        font-size: 0.8rem;
        color: #7f8c8d;
    }
    .news-date {
        font-size: 0.8rem;
        color: #95a5a6;
    }
</style>
""", unsafe_allow_html=True)

# App title and description
st.title("📰 Sector Newsletter Generator")
st.markdown("Generate a professional newsletter with the latest news and insights for your selected sectors.")

# Sidebar for inputs
with st.sidebar:
    st.header("Settings")
    
    # Sector selection
    sectors = [
        "Technology", "Finance", "Healthcare", "Energy", "Consumer Goods",
        "Industrial", "Utilities", "Real Estate", "Materials", "Communication Services",
        "Semiconductors", "Wearable Technology Sensors", "Supply Chain", "Intellectual Property Litigation"
    ]
    selected_sectors = st.multiselect(
        "Select Sectors",
        options=sectors,
        default=["Technology", "Finance"],
        help="Select one or more sectors to include in the newsletter"
    )
    
    # Newsletter settings
    st.subheader("Newsletter Settings")
    newsletter_title = st.text_input(
        "Newsletter Title",
        value="Weekly Market Insights",
        help="Customize the title of your newsletter"
    )
    
    max_articles = st.slider(
        "Max Articles per Sector",
        min_value=3,
        max_value=10,
        value=5,
        help="Maximum number of articles to include per sector"
    )
    
    # Generate button
    generate_btn = st.button("🚀 Generate Newsletter")

# Initialize news fetcher
fetcher = NewsFetcher()

def generate_wordcloud(text: str) -> str:
    """Generate a word cloud from text and return as base64"""
    wordcloud = WordCloud(
        width=800, 
        height=400, 
        background_color='white',
        max_words=100,
        contour_width=3, 
        contour_color='steelblue',
        colormap='viridis'
    ).generate(text)
    
    # Convert to base64
    img = BytesIO()
    wordcloud.to_image().save(img, format='PNG')
    return base64.b64encode(img.getvalue()).decode('utf-8')

def get_sector_icon(sector: str) -> str:
    """Get an emoji icon for a given sector"""
    icons = {
        "Semiconductors": "🔌",
        "Wearable Technology Sensors": "⌚",
        "Supply Chain": "📦",
        "Intellectual Property Litigation": "⚖️",
        # Add more mappings as needed
    }
    return icons.get(sector, "📰")  # Default to newspaper emoji

def get_sector_description(sector: str) -> str:
    """Get a description for the sector"""
    descriptions = {
        "Semiconductors": "Latest developments in semiconductor technology, chip manufacturing, and industry trends",
        "Wearable Technology Sensors": "Innovations in wearable sensors, health monitoring, and smart device technology",
        "Supply Chain": "Global supply chain updates, logistics, and industry analysis",
        "Intellectual Property Litigation": "Key patent cases, IP disputes, and legal developments in tech"
    }
    return descriptions.get(sector, f"Latest news and updates in {sector}")

def create_newsletter_presentation(news_data: Dict[str, List[Dict]], title: str) -> str:
    """Create a PowerPoint presentation from news data using the NewsletterPPTGenerator"""
    from ppt_generator.build_ppt import NewsletterPPTGenerator
    from pathlib import Path
    import os
    
    # Create output directory if it doesn't exist
    output_dir = Path('output')
    output_dir.mkdir(exist_ok=True)
    
    # Prepare configuration for the PPT generator
    config = {
        'output_dir': str(output_dir),
        'template_path': '',
        'logo_path': 'assets/logo.png'  # Update this path if you have a logo
    }
    
    # Initialize the PPT generator
    ppt_generator = NewsletterPPTGenerator(config)
    
    # Create a new presentation
    prs = ppt_generator._create_new_presentation()
    ppt_generator.prs = prs
    
    # Add title slide
    slide = ppt_generator._add_blank_slide(prs)
    
    # Add newsletter header with more spacing
    issue_date = datetime.now().strftime('%B %d, %Y')
    ppt_generator._add_newsletter_header(slide, title, f"ISSUE • {issue_date}")
    
    # Add content slides for each sector
    for sector, articles in news_data.items():
        if not articles:
            continue
            
        # Add section header slide
        slide = ppt_generator._add_blank_slide(prs)
        
        # Add section title with more spacing
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.5), 
            prs.slide_width - Inches(1), Inches(1.5)  # Increased height for better spacing
        )
        title_frame = title_box.text_frame
        title_frame.margin_bottom = Inches(0.3)  # Add bottom margin
        p = title_frame.paragraphs[0]
        p.text = sector.upper()
        p.font.size = Pt(28)  # Slightly larger font
        p.font.bold = True
        p.space_after = Pt(24)  # Add space after title
        
        # Add articles for this sector
        for article in articles:
            # Add a new slide for each article
            slide = ppt_generator._add_blank_slide(prs)
            
            # Format the article data for the card
            source = article.get('source', {})
            source_name = source.get('name', 'Unknown Source') if isinstance(source, dict) else str(source)
            
            article_data = {
                'title': article.get('title', 'No Title'),
                'source': source_name,
                'date': article.get('publishedAt', '')[:10],
                'summary': article.get('description', 'No summary available.'),
                'categories': [sector],
                'content': article.get('content', '')
            }
            
            # Add article card to the slide
            ppt_generator._add_article_card(slide, article_data, (0.5, 1.5, 12, 4.5))
    
    # Save the presentation
    output_dir = Path('output')
    output_dir.mkdir(exist_ok=True, parents=True)
    
    # Create a safe filename from the title
    safe_title = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in title.strip())
    output_path = output_dir / f"{safe_title.replace(' ', '_').lower()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
    
    try:
        prs.save(output_path)
        logger.info(f"Presentation saved to {output_path}")
        return str(output_path)
    except Exception as e:
        logger.error(f"Error saving presentation: {e}")
        raise
            


def display_news_preview(news_data: Dict[str, List[Dict]]) -> None:
    """Display a preview of the news data with proper error handling"""
    if not news_data or not isinstance(news_data, dict):
        st.warning("No news data available to display.")
        return
        
    for sector, articles in news_data.items():
        if not articles or not isinstance(articles, list):
            st.warning(f"No articles found for {sector} sector.")
            continue
            
        with st.expander(f"📰 {sector} Sector ({len(articles)} articles)", expanded=True):
            for i, article in enumerate(articles, 1):
                if not article or not isinstance(article, dict):
                    st.warning(f"Skipping invalid article #{i} in {sector} sector.")
                    continue
                    
                with st.container():
                    try:
                        # Safely get article data with defaults
                        title = str(article.get('title', 'No title')).strip() or 'Untitled Article'
                        description = str(article.get('description', '')).strip() or 'No description available'
                        description = (description[:197] + '...') if len(description) > 200 else description
                        
                        # Handle source which could be a dict or string
                        source = 'Unknown'
                        if isinstance(article.get('source'), dict):
                            source = str(article['source'].get('name', 'Unknown')).strip()
                        elif isinstance(article.get('source'), str):
                            source = article['source'].strip()
                            
                        # Format date
                        date = 'Date not available'
                        if article.get('publishedAt'):
                            try:
                                date_obj = datetime.fromisoformat(article['publishedAt'].replace('Z', '+00:00'))
                                date = date_obj.strftime('%B %d, %Y %H:%M')
                            except (ValueError, AttributeError):
                                date = str(article['publishedAt'])
                        
                        # Display article card
                        st.markdown(f"""
                        <div class="news-card" style="margin-bottom: 20px; padding: 15px; border-left: 4px solid #4CAF50; background-color: #f9f9f9;">
                            <h4 style="margin-top: 0; color: #2c3e50;">{title}</h4>
                            <p style="color: #34495e;">{description}</p>
                            <div style="font-size: 0.8em; color: #7f8c8d;">
                                <span class="news-source">Source: {source}</span>
                                <span class="news-date"> | {date}</span>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        # Add read more button if we have content
                        has_content = bool(article.get('content') or article.get('url'))
                        if has_content and st.button(f"Read more #{i}", key=f"read_more_{sector}_{i}"):
                            st.session_state[f"article_{sector}_{i}"] = not st.session_state.get(f"article_{sector}_{i}", False)
                        
                        # Show full content if expanded
                        if st.session_state.get(f"article_{sector}_{i}", False):
                            st.markdown("---")
                            st.markdown(article.get('content', 'No additional content available.'))
                            if article.get('url'):
                                st.markdown(f"[Read full article]({article['url']})")
                                
                    except Exception as e:
                        st.error(f"Error displaying article #{i} in {sector} sector: {str(e)}")
                        continue
                    
                    st.markdown("---")

def generate_newsletter():
    """Generate the newsletter based on user inputs"""
    try:
        # Reset state
        st.session_state.generated = False
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        # Step 1: Fetch news data
        update_progress(1, 5, f"Fetching latest news for {len(selected_sectors)} sectors")
        news_data = {}
        
        for i, sector in enumerate(selected_sectors, 1):
            status_text.text(f"Fetching {sector} news... ({i}/{len(selected_sectors)})")
            try:
                articles = fetcher.fetch_news(sector, max_articles)
                if not articles:
                    st.warning(f"No articles found for {sector}")
                    continue
                # Fallback logic: check sentence count and try to get better content if needed
                for idx, article in enumerate(articles):
                    content = article.get('content', '')
                    num_sentences = sum([s.strip() != '' for s in content.replace('!','.').replace('?','.').split('.')])
                    fallback_triggered = False
                    orig_len = len(content)
                    if num_sentences < 6 and article.get('url'):
                        full_content = fetcher.get_article_content(article['url'])
                        new_sentences = sum([s.strip() != '' for s in full_content.replace('!','.').replace('?','.').split('.')]) if full_content else 0
                        if full_content and len(full_content) > orig_len:
                            article['content'] = full_content
                            fallback_triggered = True
                    # Streamlit debug output for this article
                    st.write({
                        'sector': sector,
                        'article_idx': idx,
                        'title': article.get('title',''),
                        'orig_sentence_count': num_sentences,
                        'fallback_triggered': fallback_triggered,
                        'new_sentence_count': sum([s.strip() != '' for s in article.get('content','').replace('!','.').replace('?','.').split('.')])
                    })
                news_data[sector] = articles
                st.success(f"Found {len(articles)} articles for {sector}")
            except Exception as e:
                st.error(f"Error fetching {sector} news: {str(e)}")
                continue
        
        st.session_state.news_data = news_data
        
        # Step 2: Process and analyze data
        update_progress(2, 5, "Processing and analyzing data")
        # TODO: Add more sophisticated analysis here
        
        # Step 3: Generate visualizations
        update_progress(3, 5, "Generating visualizations")
        # Visualizations are generated on-demand in the preview
        
        # Step 4: Create PowerPoint
        update_progress(4, 5, "Creating PowerPoint presentation")
        output_path = create_newsletter_presentation(news_data, newsletter_title)
        
        # Update state
        st.session_state.generated = True
        st.session_state.newsletter_path = output_path
        
        update_progress(5, 5, "Done!")
        status_text.success("Newsletter generated successfully!")
        
    except Exception as e:
        status_text.error(f"Error generating newsletter: {str(e)}")
        st.exception(e)

# Main content area
newsletter_placeholder = st.empty()
status_text = st.empty()
progress_bar = st.progress(0)

def update_progress(step: int, total_steps: int, message: str) -> None:
    """Update progress bar and status text"""
    progress = (step / total_steps) * 100
    progress_bar.progress(int(progress))
    status_text.text(f"Status: {message}...")

# Handle generate button click
if generate_btn and selected_sectors:
    with newsletter_placeholder:
        st.info("Generating your newsletter. This may take a few minutes...")
        generate_newsletter()

# Display preview if data is available
if st.session_state.news_data:
    st.subheader("Newsletter Preview")
    display_news_preview(st.session_state.news_data)

# Show download button if newsletter is generated
if st.session_state.get('generated', False):
    st.download_button(
        label="📥 Download Newsletter (PowerPoint)",
        data=open(st.session_state.newsletter_path, 'rb'),
        file_name=f"{newsletter_title.replace(' ', '_').lower()}_{datetime.now().strftime('%Y%m%d')}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

# Add some helpful tips
with st.expander("💡 Tips for Better Results"):
    st.markdown("""
    - Select 2-3 sectors for a focused newsletter
    - Use specific keywords for more targeted results
    - Check back regularly for the latest updates
    - Customize the newsletter title to match your needs
    - The system caches results to avoid redundant API calls
    """)
