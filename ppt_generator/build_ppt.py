from typing import List, Dict, Any, Optional, Tuple
from pathlib import Path
from datetime import datetime
import logging
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
import os
import json
import random

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class NewsletterPPTGenerator:
    """Generate newsletter-style PowerPoint presentations from article insights"""
    
    def __init__(self, config: dict):
        """Initialize the PPT generator with configuration"""
        self.config = config
        self.template_path = Path(config.get('template_path', ''))
        self.output_dir = Path(config.get('output_dir', 'output'))
        self.logo_path = Path(config.get('logo_path', 'assets/logo.png'))
        
        # Create output directory if it doesn't exist
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Professional color palette
        self.colors = {
            'primary': (13, 59, 102),      # Dark blue
            'secondary': (27, 85, 226),    # Vibrant blue
            'accent': (1, 180, 228),       # Cyan accent
            'success': (57, 181, 74),      # Fresh green
            'warning': (255, 183, 3),      # Amber
            'background': (250, 250, 252), # Off-white
            'white': (255, 255, 255),      # Pure white
            'dark': (20, 24, 36),          # Near black
            'light_text': (108, 117, 125), # Gray text
            'border': (233, 236, 239),     # Light border
            'gradient_start': (245, 247, 250),  # Light gray gradient start
            'gradient_end': (255, 255, 255)     # White gradient end
        }
        
        # Typography settings with fallbacks for cross-platform compatibility
        self.fonts = {
            'heading': 'Montserrat, Arial, sans-serif',
            'subheading': 'Open Sans, Arial, sans-serif',
            'body': 'Segoe UI, Arial, sans-serif',
            'accent': 'Georgia, Times New Roman, serif',
            'mono': 'Courier New, monospace'
        }
        
        # Slide layouts
        self.slide_layouts = {
            'title': 0,
            'section_header': 1,
            'content': 2,
            'two_content': 3,
            'comparison': 4,
            'title_only': 5,
            'blank': 6,
            'content_with_caption': 7,
            'picture_with_caption': 8
        }
    
    def _create_new_presentation(self):
        """Create a new presentation with default settings"""
        from pptx import Presentation
        from pptx.util import Inches
        
        # Create a new presentation with a blank layout
        prs = Presentation()
        
        # Set slide size to 16:9 (widescreen)
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        
        # Return the presentation
        return prs
        
    def _add_blank_slide(self, prs):
        """Add a blank slide to the presentation"""
        # Use the blank layout (index 6)
        blank_slide_layout = prs.slide_layouts[6]
        return prs.slides.add_slide(blank_slide_layout)
        
    def _add_newsletter_header(self, slide, title: str, subtitle: str = "") -> None:
        """Add a professional newsletter header to the slide"""
        # Add a rectangle for the header background
        header = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            0, 0,
            slide.width, Inches(1.5)
        )
        
        # Style the header
        fill = header.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*self.colors['primary'])
        header.line.fill.background()
        
        # Add logo if available
        if hasattr(self, 'logo_path') and self.logo_path.exists():
            try:
                logo = slide.shapes.add_picture(
                    str(self.logo_path),
                    Inches(0.5), Inches(0.25),
                    height=Inches(1.0)
                )
                # Position title to the right of the logo
                title_left = logo.left + logo.width + Inches(0.5)
            except Exception as e:
                logger.warning(f"Could not add logo: {e}")
                title_left = Inches(0.5)
        else:
            title_left = Inches(0.5)
        
        # Add title
        title_box = slide.shapes.add_textbox(
            title_left, Inches(0.4),
            slide.width - title_left - Inches(0.5), Inches(0.7)
        )
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = title.upper()
        p.font.name = self.fonts['heading'].split(',')[0].strip()
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = RGBColor(*self.colors['white'])
        p.space_after = 0
        
        # Add subtitle if provided
        if subtitle:
            subtitle_box = slide.shapes.add_textbox(
                title_left, Inches(0.9),
                slide.width - title_left - Inches(0.5), Inches(0.4)
            )
            tf = subtitle_box.text_frame
            p = tf.paragraphs[0]
            p.text = subtitle.upper()
            p.font.name = self.fonts['subheading'].split(',')[0].strip()
            p.font.size = Pt(12)
            p.font.color.rgb = RGBColor(*self.colors['white'])
            p.space_after = 0
        
        # Add a subtle separator line
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            0, Inches(1.5) - Pt(2),
            slide.width, Pt(2)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(*self.colors['accent'])
        line.line.fill.background()
        
    def _create_gradient_background(self, slide) -> None:
        """Create a subtle gradient background"""
        # Add a rectangle shape for background
        left = 0
        top = 0
        width = slide.parent.slide_width
        height = slide.parent.slide_height
        
        bg_shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, left, top, width, height
        )
        
        # Set fill to gradient
        fill = bg_shape.fill
        fill.gradient()
        fill.gradient_angle = 45
        
        # Gradient stops
        fill.gradient_stops[0].position = 0.0
        fill.gradient_stops[0].color.rgb = RGBColor(*self.colors['gradient_start'])
        fill.gradient_stops[1].position = 1.0
        fill.gradient_stops[1].color.rgb = RGBColor(*self.colors['gradient_end'])
        
        # Remove border
        bg_shape.line.fill.background()
        
        # Send to back
        bg_shape.z_order = 0
    
    def _add_decorative_header_bar(self, slide) -> None:
        """Add a modern header bar with subtle shadow"""
        # Main header bar
        left = 0
        top = 0
        width = slide.parent.slide_width
        height = Inches(0.25)
        
        # Add subtle shadow effect (semi-transparent black)
        shadow = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, 
            left + Pt(2), top + Pt(2), 
            width, height
        )
        shadow.fill.solid()
        shadow.fill.fore_color.rgb = RGBColor(0, 0, 0)
        shadow.fill.alpha = 0.1
        shadow.line.fill.background()
        
        # Main header bar
        header_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, left, top, width, height
        )
        
        # Gradient fill
        fill = header_bar.fill
        fill.gradient()
        fill.gradient_angle = 0
        
        # Gradient from primary to secondary color
        fill.gradient_stops[0].position = 0.0
        fill.gradient_stops[0].color.rgb = RGBColor(*self.colors['primary'])
        fill.gradient_stops[1].position = 1.0
        fill.gradient_stops[1].color.rgb = RGBColor(*self.colors['secondary'])
        
        # Remove border
        header_bar.line.fill.background()
        
        # Add a thin accent line at the bottom
        accent_line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left, top + height - Pt(2),
            width, Pt(4)
        )
        accent_line.fill.solid()
        accent_line.fill.fore_color.rgb = RGBColor(*self.colors['accent'])
        accent_line.line.fill.background()
    
    def _add_newsletter_header(self, slide, title: str, issue_info: str = "") -> None:
        """Add a clean, structured newsletter header"""
        # Create a header box
        header_box = self._create_content_box(
            slide, 
            (0.25, 0.25, 13.5, 1.5),
            has_border=False
        )
        
        # Add logo if available
        if self.logo_path.exists():
            try:
                logo_left = Inches(0.5)
                logo_top = Inches(0.4)
                logo = slide.shapes.add_picture(
                    str(self.logo_path), 
                    logo_left, logo_top, 
                    height=Inches(0.7)
                )
                title_left = Inches(1.8)
            except Exception as e:
                logger.warning(f"Could not add logo: {e}")
                title_left = Inches(0.5)
        else:
            title_left = Inches(0.5)
        
        # Add title
        title_box = slide.shapes.add_textbox(
            Inches(title_left), Inches(0.4),
            Inches(8), Inches(0.8)
        )
        title_frame = title_box.text_frame
        
        # Main title
        p = title_frame.paragraphs[0]
        p.text = title.upper()
        p.font.name = self.fonts['heading'].split(',')[0].strip()
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = RGBColor(*self.colors['primary'])
        p.space_after = Pt(4)
        
        # Add issue info
        if issue_info:
            p = title_frame.add_paragraph()
            p.text = issue_info.upper()
            p.font.name = self.fonts['subheading'].split(',')[0].strip()
            p.font.size = Pt(10)
            p.font.color.rgb = RGBColor(*self.colors['light_text'])
            p.space_before = 0
            p.space_after = 0
        
        # Add date on the right side
        if issue_info:
            date_box = slide.shapes.add_textbox(
                Inches(10), Inches(0.6),
                Inches(3), Inches(0.5)
            )
            date_frame = date_box.text_frame
            p = date_frame.paragraphs[0]
            p.text = datetime.now().strftime('%B %d, %Y')
            p.font.name = self.fonts['body'].split(',')[0].strip()
            p.font.size = Pt(10)
            p.font.italic = True
            p.font.color.rgb = RGBColor(*self.colors['light_text'])
            p.alignment = PP_ALIGN.RIGHT
    
    def _create_content_box(self, slide, position: Tuple[float, float, float, float], 
                          title: str = "", has_border: bool = True) -> tuple:
        """Create a structured content box with optional title"""
        left, top, width, height = position
        
        # Convert to points for precision
        left_pt = Inches(left)
        top_pt = Inches(top)
        width_pt = Inches(width)
        height_pt = Inches(height)
        
        # Create main box
        box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left_pt, top_pt,
            width_pt, height_pt
        )
        
        # Style the box
        fill = box.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*self.colors['white'])
        
        line = box.line
        if has_border:
            line.color.rgb = RGBColor(*self.colors['border'])
            line.width = Pt(1)
        else:
            line.fill.background()
        
        # Add title if provided
        title_box = None
        if title:
            # Title background
            title_bg = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                left_pt, top_pt,
                width_pt, Pt(40)
            )
            title_bg.fill.solid()
            title_bg.fill.fore_color.rgb = RGBColor(*self.colors['primary'])
            title_bg.line.fill.background()
            
            # Title text
            title_box = slide.shapes.add_textbox(
                left_pt + Pt(10), top_pt + Pt(5),
                width_pt - Pt(20), Pt(30)
            )
            tf = title_box.text_frame
            p = tf.paragraphs[0]
            p.text = title.upper()
            p.font.name = self.fonts['heading'].split(',')[0].strip()
            p.font.size = Pt(14)
            p.font.bold = True
            p.font.color.rgb = RGBColor(*self.colors['white'])
            p.space_after = 0
            
            # Adjust content area
            top_pt += Pt(45)
            height_pt -= Pt(45)
        
        # Return content area coordinates
        return (left_pt, top_pt, width_pt, height_pt, box, title_box)
        
    def _add_article_card(self, slide, article: Dict[str, Any], 
                         position: Tuple[float, float, float, float]) -> None:
        """Add a visually appealing article card with enhanced content"""
        from pptx.util import Inches, Pt
        from textwrap import wrap
        
        # Convert position tuple to inches
        left, top, width, height = position
        left_pt = Inches(left)
        top_pt = Inches(top)
        width_pt = Inches(width)
        height_pt = Inches(height)
        
        # Add a card with shadow
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            left_pt, top_pt, width_pt, height_pt
        )
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(255, 255, 255)  # White background
        card.line.color.rgb = RGBColor(230, 230, 230)  # Light gray border
        
        # Add shadow effect (using default shadow color)
        shadow = card.shadow
        shadow.inherit = False
        shadow.visible = True
        shadow.blur_radius = Pt(4)
        shadow.offset_x = Pt(2)
        shadow.offset_y = Pt(2)
        # Add subtle accent line at top
        accent = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left_pt, top_pt,
            width_pt, Pt(4)
        )
        accent.fill.solid()
        accent.fill.fore_color.rgb = RGBColor(1, 180, 228)  # Accent color
        accent.line.fill.background()
        
        # Add content with padding
        padding = Pt(12)
        content_left = left_pt + padding
        content_top = top_pt + padding + Pt(8)  # Extra space for accent
        content_width = width_pt - (2 * padding)
        
        # Add title with better typography
        title = article.get('title', 'No Title')
        title_box = slide.shapes.add_textbox(
            content_left, content_top,
            content_width, Pt(40)
        )
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = 'Arial'
        p.font.size = Pt(16)  # Slightly larger for better readability
        p.font.bold = True
        p.font.color.rgb = RGBColor(20, 30, 40)
        p.space_after = Pt(6)
        
        # Add source and date with better styling
        source = article.get('source', 'Source')
        date = article.get('date', datetime.now().strftime('%b %d, %Y'))
        source_text = f"{source.upper()} • {date}"
        
        source_box = slide.shapes.add_textbox(
            content_left, content_top + Pt(28),
            content_width, Pt(15)
        )
        tf = source_box.text_frame
        p = tf.paragraphs[0]
        p.text = source_text
        p.font.name = 'Arial'
        p.font.size = Pt(9)
        p.font.color.rgb = RGBColor(1, 180, 228)  # Accent color
        p.space_after = Pt(12)
        
        # Generate more elaborate summary if needed
        summary = article.get('summary', 'No summary available.')
        if len(summary) < 200:  # If summary is too short, enhance it
            summary = self._enhance_summary(summary, article)
            
        # Add summary with better formatting
        summary_box = slide.shapes.add_textbox(
            content_left, content_top + Pt(50),
            content_width, height_pt - Pt(90)  # More space for content
        )
        tf = summary_box.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        
        # Split summary into paragraphs for better readability
        paragraphs = summary.split('\n\n')
        for i, para in enumerate(paragraphs):
            if i > 0:
                p = tf.add_paragraph()
                p.text = ''  # Add space between paragraphs
                p.space_after = Pt(6)
                
            p = tf.add_paragraph()
            p.text = para
            p.font.name = 'Arial'
            p.font.size = Pt(10)
            p.space_after = Pt(6)
        
        # Add category tags at bottom if available
        if 'categories' in article and article['categories']:
            categories = article['categories'][:2]  # Show max 2 categories
            tag_top = top_pt + height_pt - Pt(30)
            
            for i, category in enumerate(categories):
                if i > 1:  # Only show first 2 categories
                    break
                    
                # Create tag background
                tag_width = min(Inches(1.5), Inches(len(category) * 0.15))
                tag_left = content_left + (Inches(1.8) * i)
                
                tag_bg = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    tag_left, tag_top,
                    tag_width, Pt(18)
                )
                tag_bg.fill.solid()
                tag_bg.fill.fore_color.rgb = RGBColor(240, 245, 250)  # Light blue-gray
                tag_bg.line.color.rgb = RGBColor(200, 220, 240)
                tag_bg.line.width = Pt(0.5)
                
                # Add tag text
                tag_text = slide.shapes.add_textbox(
                    tag_left + Pt(4), tag_top,
                    tag_width - Pt(8), Pt(18)
                )
                tf = tag_text.text_frame
                p = tf.paragraphs[0]
                p.text = category.upper()
                p.font.name = 'Arial'
                p.font.size = Pt(8)
                p.font.bold = True
                p.font.color.rgb = RGBColor(70, 100, 150)
                p.alignment = PP_ALIGN.CENTER
    
    def _enhance_summary(self, summary: str, article: Dict[str, Any]) -> str:
        """Enhance the summary with more details if it's too short"""
        if not summary:
            summary = article.get('content', '')[:500] + '...' if article.get('content') else 'No summary available.'
            
        if len(summary) > 200:  # Already long enough
            return summary
            
        # Add key points if available
        key_points = []
        if 'key_points' in article and article['key_points']:
            key_points = article['key_points']
        
        # If no key points, generate some from the content
        if not key_points and 'content' in article and article['content']:
            content = article['content']
            # Simple way to extract key sentences (basic implementation)
            sentences = [s.strip() for s in content.split('.') if len(s.split()) > 5]
            if len(sentences) > 3:
                key_points = [s + '.' for s in sentences[:3]]
        
        # Enhance the summary
        enhanced = summary
        if key_points:
            enhanced += "\n\nKey Points:"
            for point in key_points:
                enhanced += f"\n• {point}"
        
        # Add source and date if not already in summary
        source = article.get('source', '')
        date = article.get('date', '')
        if source and source not in enhanced:
            enhanced += f"\n\nSource: {source}"
        if date and date not in enhanced:
            enhanced += f" | {date}"
            
        return enhanced
    
    def _add_stats_visualization(self, slide, insights: List[Dict[str, Any]]) -> None:
        """Add a simple statistics visualization"""
        if not insights:
            return
            
        # Count insights by category
        category_counts = {}
        for insight in insights:
            for category in insight.get('categories', ['Other']):
                category_counts[category] = category_counts.get(category, 0) + 1
        
        # Create a simple bar chart representation
        chart_left = Inches(1)
        chart_top = Inches(2)
        chart_width = Inches(8)
        chart_height = Inches(3)
        
        # Chart title
        title_box = slide.shapes.add_textbox(
            chart_left, Inches(1.5), chart_width, Inches(0.4)
        )
        title_frame = title_box.text_frame
        p = title_frame.paragraphs[0]
        p.text = "Articles by Category"
        p.font.name = 'Arial'
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = RGBColor(*self.colors['primary'])
        p.alignment = PP_ALIGN.CENTER

        # Create bars
        max_count = max(category_counts.values()) if category_counts else 1
        bar_height = 0.4
        bar_spacing = 0.6
        colors = [self.colors['primary'], self.colors['secondary'], self.colors['accent'], 
                 self.colors['success'], self.colors['warning']]

        y_pos = chart_top
        for i, (category, count) in enumerate(category_counts.items()):
            # Bar
            bar_width = (count / max_count) * (chart_width.inches - 2)
            
            # Create bar shape
            bar = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(chart_left.inches + 2), 
                y_pos,
                Inches(bar_width), 
                Inches(bar_height)
            )
            
            # Style the bar
            color = colors[i % len(colors)]
            bar.fill.solid()
            bar.fill.fore_color.rgb = RGBColor(*color)
            bar.line.fill.background()
            
            # Add category label
            label = slide.shapes.add_textbox(
                chart_left, 
                y_pos,
                Inches(1.8), 
                Inches(bar_height)
            )
            
            # Configure label text
            tf = label.text_frame
            p = tf.paragraphs[0]
            p.text = f"{category} ({count})"
            p.font.name = 'Arial'
            p.font.size = Pt(10)
            p.font.color.rgb = RGBColor(*self.colors['dark'])
            p.alignment = PP_ALIGN.LEFT
            
            # Position for next bar
            y_pos += Inches(bar_height + 0.2)
            
            # Stop if we're running out of vertical space
            if y_pos > Inches(6.5):
                break
                
        # Add a simple legend
        legend_left = Inches(8)
        legend_top = Inches(2)
        legend_size = Inches(0.3)
        
        for i, (category, _) in enumerate(category_counts.items()):
            if i >= 5:  # Limit legend items
                break
                
            # Color swatch
            swatch = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                legend_left,
                legend_top + (i * Inches(0.5)),
                legend_size,
                legend_size
            )
            swatch.fill.solid()
            swatch.fill.fore_color.rgb = RGBColor(*colors[i % len(colors)])
            swatch.line.fill.background()
            
            # Category name
            cat_label = slide.shapes.add_textbox(
                legend_left + legend_size + Inches(0.2),
                legend_top + (i * Inches(0.5)),
                Inches(2),
                legend_size
            )
            tf = cat_label.text_frame
            p = tf.paragraphs[0]
            p.text = category
            p.font.name = 'Arial'
            p.font.size = Pt(9)
            p.font.color.rgb = RGBColor(*self.colors['dark'])
            
        # Add a call to action at the bottom
        cta_box = slide.shapes.add_textbox(
            Inches(1), Inches(6),
            slide.width - Inches(2), Inches(1)
        )
        tf = cta_box.text_frame
        p = tf.paragraphs[0]
        p.text = "Read more about these articles and stay up-to-date with the latest news!"
        p.font.name = 'Arial'
        p.font.size = Pt(14)
        p.font.color.rgb = RGBColor(100, 100, 100)
        p.alignment = PP_ALIGN.CENTER

    def generate_presentation(self, insights: List[Dict[str, Any]], 
                            output_path: Optional[str] = None) -> str:
        """
{{ ... }}
        
        Args:
            insights: List of insight dictionaries
            output_path: Path to save the presentation (optional)
            
        Returns:
            Path to the generated presentation
        """
        if not insights:
            raise ValueError("No insights provided to generate presentation")
        
        # Create a new presentation
        if self.template_path and self.template_path.exists():
            prs = Presentation(str(self.template_path))
        else:
            prs = Presentation()
        
        # Set presentation properties (16:9 aspect ratio)
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)
        
        # Add title slide
        sector = insights[0].get('sector', 'Industry')
        self._add_title_slide(
            prs, 
            title=f"{sector} Intelligence Newsletter",
            subtitle="Monthly Insights & Market Analysis"
        )
        
        # Group insights by category
        categorized = {}
        for insight in insights:
            cats = insight.get('categories', ['General'])
            for cat in cats:
                if cat not in categorized:
                    categorized[cat] = []
                categorized[cat].append(insight)
        
        # Add content slides for each category
        for category, items in categorized.items():
            # Add section header
            self._add_section_header(prs, category)
            
            # Add content slides (max 3 articles per slide for readability)
            for i in range(0, len(items), 3):
                batch = items[i:i+3]
                self._add_newsletter_content_slide(prs, batch, category)
        
        # Add summary slide
        self._add_summary_slide(prs, insights)
        
        # Save the presentation
        if not output_path:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = self.output_dir / f"{sector.lower().replace(' ', '_')}_newsletter_{timestamp}.pptx"
        else:
            output_path = Path(output_path)
        
        prs.save(str(output_path))
        logger.info(f"Newsletter presentation saved to {output_path}")
        
        return str(output_path)


def load_insights_from_file(file_path: str) -> List[Dict[str, Any]]:
    """
    Load insights from a JSON file
    
    Args:
        file_path: Path to JSON file containing insights
        
    Returns:
        List of insight dictionaries
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        logger.error(f"Error loading insights from file: {str(e)}")
        return []


def generate_sample_insights() -> List[Dict[str, Any]]:
    """Generate sample insights for testing the newsletter format"""
    return [
        {
            "title": "AI Revolution Transforms Healthcare Diagnostics",
            "content": "Revolutionary AI technology is enabling faster and more accurate medical diagnoses, with new algorithms showing 95% accuracy in detecting early-stage diseases. The technology combines machine learning with advanced imaging to provide real-time analysis that could save millions of lives annually.",
            "source": "MedTech Today",
            "date": "2024-06-15",
            "categories": ["Healthcare", "Artificial Intelligence", "Innovation"],
            "sector": "Healthcare Technology",
            "key_points": [
                "95% accuracy in early disease detection",
                "Real-time analysis capabilities",
                "Potential to save millions of lives"
            ]
        },
        {
            "title": "Green Energy Breakthrough: Solar Efficiency Reaches New Heights",
            "content": "Scientists have achieved a major breakthrough in solar panel technology, reaching 47% efficiency in laboratory conditions. This advancement could revolutionize renewable energy adoption and significantly reduce costs for consumers worldwide.",
            "source": "Energy Innovation Weekly",
            "date": "2024-06-12",
            "categories": ["Renewable Energy", "Technology", "Sustainability"],
            "sector": "Clean Energy",
            "key_points": [
                "47% solar panel efficiency achieved",
                "Could revolutionize renewable energy",
                "Significant cost reduction potential"
            ]
        },
        {
            "title": "Quantum Computing Makes Commercial Debut",
            "content": "The first commercial quantum computer has been deployed for financial modeling, promising to solve complex optimization problems in seconds rather than hours. This marks a significant milestone in quantum technology commercialization.",
            "source": "Quantum Business Review",
            "date": "2024-06-10",
            "categories": ["Quantum Computing", "Finance", "Technology"],
            "sector": "Quantum Technology",
            "key_points": [
                "First commercial quantum deployment",
                "Solves complex problems in seconds",
                "Major milestone for commercialization"
            ]
        },
        {
            "title": "Space Tourism Industry Reaches New Milestone",
            "content": "Commercial space flights have successfully transported over 1,000 passengers to the edge of space, marking a significant achievement for the space tourism industry. The milestone demonstrates growing consumer confidence and technological maturity.",
            "source": "Space Commerce Daily",
            "date": "2024-06-08",
            "categories": ["Space Tourism", "Transportation", "Innovation"],
            "sector": "Aerospace",
            "key_points": [
                "Over 1,000 passengers transported",
                "Growing consumer confidence",
                "Technological maturity demonstrated"
            ]
        },
        {
            "title": "Autonomous Vehicles Pass Major Safety Milestone",
            "content": "Self-driving cars have achieved a 99.9% safety record in controlled testing environments, surpassing human driver performance in multiple categories. The breakthrough brings fully autonomous transportation closer to reality.",
            "source": "AutoTech Insider",
            "date": "2024-06-05",
            "categories": ["Autonomous Vehicles", "Safety", "Transportation"],
            "sector": "Automotive Technology",
            "key_points": [
                "99.9% safety record achieved",
                "Surpasses human driver performance",
                "Brings full autonomy closer to reality"
            ]
        }
    ]


# Example usage
if __name__ == "__main__":
    # Configuration
    config = {
        'output_dir': 'newsletter_output',
        'logo_path': 'assets/company_logo.png'
    }
    
    # Create generator
    generator = NewsletterPPTGenerator(config)
    
    # Generate sample presentation
    sample_insights = generate_sample_insights()
    output_file = generator.generate_presentation(sample_insights)
    
    print(f"Newsletter presentation generated: {output_file}")