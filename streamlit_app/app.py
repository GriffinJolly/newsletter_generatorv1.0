import streamlit as st
import os
from datetime import datetime, timedelta
import yaml
from pathlib import Path
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
    """Create a PowerPoint presentation from news data"""
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
    
    prs = Presentation()
    
    # Set slide size to 16:9
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    # Title slide
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    
    # Format title
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.size = Pt(44)
    title_shape.text_frame.paragraphs[0].font.bold = True
    
    # Add subtitle with date and sectors
    subtitle = slide.placeholders[1]
    subtitle.text = f"{get_sector_icon('')} Generated on {datetime.now().strftime('%B %d, %Y')}"
    subtitle.text_frame.paragraphs[0].font.size = Pt(18)
    
    # Add a text box for sectors
    left = Inches(1)
    top = Inches(3)
    width = Inches(11.3)
    height = Inches(2)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    
    p = tf.add_paragraph()
    p.text = "Covering Sectors:"
    p.font.size = Pt(14)
    p.font.bold = True
    
    # Add sectors with icons
    p = tf.add_paragraph()
    for sector in selected_sectors:
        p.text += f"{get_sector_icon(sector)} {sector}  "
    p.font.size = Pt(16)
    p.space_after = Pt(12)
    
    # Add a summary slide with improved formatting
    slide_layout = prs.slide_layouts[5]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add title
    left = Inches(0.5)
    top = Inches(0.5)
    width = Inches(12)
    height = Inches(1)
    title_box = slide.shapes.add_textbox(left, top, width, height)
    title_frame = title_box.text_frame
    title_p = title_frame.add_paragraph()
    title_p.text = "EXECUTIVE SUMMARY"
    title_p.font.size = Pt(28)
    title_p.font.bold = True
    title_p.font.color.rgb = RGBColor(59, 89, 152)  # Dark blue
    
    # Add a divider line
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, top + Inches(0.6), width, Inches(0.1)
    )
    fill = line.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(59, 89, 152)
    line.line.fill.background()
    
    # Add summary content
    left = Inches(1)
    top = Inches(1.5)
    width = Inches(11)
    height = Inches(5)
    content_box = slide.shapes.add_textbox(left, top, width, height)
    content_frame = content_box.text_frame
    
    # Add summary points
    p = content_frame.add_paragraph()
    p.text = f"📊 Market Overview"
    p.font.size = Pt(20)
    p.font.bold = True
    p.space_after = Pt(12)
    
    p = content_frame.add_paragraph()
    p.text = f"• Comprehensive analysis of {len(selected_sectors)} key sectors"
    p.font.size = Pt(14)
    p.level = 0
    
    p = content_frame.add_paragraph()
    p.text = f"• {sum(len(articles) for articles in news_data.values())} articles analyzed"
    p.font.size = Pt(14)
    p.level = 0
    
    # Add sector highlights
    p = content_frame.add_paragraph()
    p.text = "\n🔍 Sector Highlights"
    p.font.size = Pt(20)
    p.font.bold = True
    p.space_before = Pt(20)
    p.space_after = Pt(12)
    
    for sector in selected_sectors:
        p = content_frame.add_paragraph()
        p.text = f"{get_sector_icon(sector)} {sector}:"
        p.font.size = Pt(14)
        p.font.bold = True
        p.level = 0
        
        p = content_frame.add_paragraph()
        p.text = get_sector_description(sector)
        p.font.size = Pt(12)
        p.level = 1
        p.space_after = Pt(8)
    
    # Add a slide for each sector
    for sector, articles in news_data.items():
        if not articles:
            logger.warning(f"No articles found for {sector}, skipping...")
            continue
            
        try:
            # Sector title slide
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            title = slide.shapes.title
            content = slide.placeholders[1]
            title.text = f"{sector} Sector Update"
            
            # Add a word cloud for the sector
            all_text = ' '.join([
                str(a.get('content', '') or a.get('description', '') or a.get('title', '') or '')
                for a in articles
            ]).strip()
            
            if all_text:
                try:
                    img_data = generate_wordcloud(all_text)
                    img_path = f"wordcloud_{sector.lower().replace(' ', '_')}.png"
                    with open(img_path, 'wb') as f:
                        f.write(base64.b64decode(img_data))
                    
                    slide = prs.slides.add_slide(prs.slide_layouts[5])
                    slide.shapes.title.text = f"{sector} - Word Cloud"
                    left = Inches(1)
                    top = Inches(1.5)
                    height = Inches(5)
                    slide.shapes.add_picture(img_path, left, top, height=height)
                    
                    try:
                        os.remove(img_path)
                    except Exception as e:
                        logger.warning(f"Could not remove temporary file {img_path}: {e}")
                        
                except Exception as e:
                    logger.error(f"Error generating word cloud for {sector}: {e}")
                    
            # Add slides for each article
            for article in articles:
                try:
                    slide = prs.slides.add_slide(prs.slide_layouts[1])
                    title = slide.shapes.title
                    content = slide.placeholders[1]
                    
                    title.text = article.get('title', 'No Title')
                    
                    # Format the content
                    source = article.get('source', {}) if isinstance(article.get('source'), dict) else {'name': str(article.get('source', 'Unknown'))}
                    published = article.get('publishedAt', '')
                    if published:
                        try:
                            published = datetime.strptime(published, '%Y-%m-%dT%H:%M:%SZ').strftime('%B %d, %Y')
                        except:
                            pass
                    
                    content.text = f"""
                    Source: {source.get('name', 'Unknown')}
                    Published: {published}
                    
                    {article.get('description', 'No description available')}
                    
                    {article.get('content', '')}
                    """.strip()
                    
                except Exception as e:
                    logger.error(f"Error creating slide for article: {e}")
                    continue
                    
        except Exception as e:
            logger.error(f"Error processing sector {sector}: {e}")
            continue
            
    # Add a summary slide at the end
    try:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        title = slide.shapes.title
        content = slide.placeholders[1]
        title.text = "Key Insights & Next Steps"
        content.text = "• Review the latest market trends\n• Consider the impact on your portfolio\n• Stay tuned for our next update"
    except Exception as e:
        logger.error(f"Error creating summary slide: {e}")
        
    # Save the presentation
    try:
        title_str = str(title).strip() if title is not None else "newsletter"
        safe_title = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in title_str)
        newsletter_filename = f"{safe_title.replace(' ', '_').lower()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        output_path = str(Path(CONFIG['output']['directory']) / newsletter_filename)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        prs.save(output_path)
        logger.info(f"Presentation saved to {output_path}")
        return output_path
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
