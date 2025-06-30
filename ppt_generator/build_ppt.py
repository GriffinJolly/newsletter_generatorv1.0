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
        """Add a simple and reliable article card"""
        from pptx.util import Inches, Pt
        
        # Convert position tuple to inches
        left, top, width, height = position
        left_pt = Inches(left)
        top_pt = Inches(top)
        width_pt = Inches(width)
        height_pt = Inches(height)
        
        # Add a simple white background
        box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left_pt, top_pt, width_pt, height_pt
        )
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        box.line.color.rgb = RGBColor(220, 220, 220)
        box.line.width = Pt(0.5)
        
        # Add content with padding
        padding = Pt(10)
        content_left = left_pt + padding
        content_top = top_pt + padding
        content_width = width_pt - (2 * padding)
        
        # Add title
        title_box = slide.shapes.add_textbox(
            content_left, content_top,
            content_width, Pt(40)
        )
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = article.get('title', 'No Title')
        p.font.name = 'Arial'
        p.font.size = Pt(14)
        p.font.bold = True
        p.space_after = Pt(8)
        
        # Add source and date below title
        source_text = f"{article.get('source', 'Source')} • {article.get('date', '')}"
        source_box = slide.shapes.add_textbox(
            content_left, content_top + Pt(25),
            content_width, Pt(15)
        )
        tf = source_box.text_frame
        p = tf.paragraphs[0]
        p.text = source_text.upper()
        p.font.name = 'Arial'
        p.font.size = Pt(8)
        p.font.color.rgb = RGBColor(100, 100, 100)
        p.space_after = Pt(15)
        
        # Add summary
        summary_box = slide.shapes.add_textbox(
            content_left, content_top + Pt(50),
            content_width, height_pt - Pt(80)
        )
        tf = summary_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = article.get('summary', 'No summary available.')
        p.font.name = 'Arial'
        p.font.size = Pt(10)
        p.space_after = Pt(6)
        
        # Add category tag at bottom if available
        if 'categories' in article and article['categories']:
            category = article['categories'][0][:20]  # Limit length
            tag_top = top_pt + height_pt - Pt(25)
            
            # Simple text tag
            tag_box = slide.shapes.add_textbox(
                content_left, tag_top,
                Pt(100), Pt(15)
            )
            tf = tag_box.text_frame
            p = tf.paragraphs[0]
            p.text = f"{category.upper()}"
            p.font.name = 'Arial'
            p.font.size = Pt(8)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0, 120, 212)  # Blue accent
    
    def _add_stats_visualization(self, slide, insights: List[Dict[str, Any]]) -> None:
        """Add a simple statistics visualization"""
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
        p.font.name = self.fonts['subheading']
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
            bar = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(chart_left.inches + 2), Inches(y_pos.inches),
                Inches(bar_width), Inches(bar_height)
            )
            
            color = colors[i % len(colors)]
            bar.fill.solid()
            bar.fill.fore_color.rgb = RGBColor(*color)
            bar.line.fill.background()
            
            # Category label
            label_box = slide.shapes.add_textbox(
                chart_left, Inches(y_pos.inches),
                Inches(1.8), Inches(bar_height)
            )
            label_frame = label_box.text_frame
            p = label_frame.paragraphs[0]
            p.text = category
            p.font.name = self.fonts['body']
            p.font.size = Pt(10)
            p.font.color.rgb = RGBColor(*self.colors['dark'])
            p.alignment = PP_ALIGN.RIGHT
            
            # Count label
            count_box = slide.shapes.add_textbox(
                Inches(chart_left.inches + 2 + bar_width + 0.1),
                Inches(y_pos.inches),
                Inches(0.5), Inches(bar_height)
            )
            count_frame = count_box.text_frame
            p = count_frame.paragraphs[0]
            p.text = str(count)
            p.font.name = self.fonts['body']
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = RGBColor(*color)
            
            y_pos = Inches(y_pos.inches + bar_spacing)
    
    def _add_title_slide(self, prs: Presentation, title: str, subtitle: str = "") -> None:
        """Add a newsletter-style title slide"""
        slide_layout = prs.slide_layouts[self.slide_layouts['blank']]
        slide = prs.slides.add_slide(slide_layout)
        
        # Create gradient background
        self._create_gradient_background(slide, self.colors['background'], self.colors['white'])
        
        # Add decorative elements
        self._add_decorative_header_bar(slide, self.colors['primary'])
        
        # Main title
        left = Inches(1)
        top = Inches(2)
        width = Inches(10)
        height = Inches(2)
        
        title_box = slide.shapes.add_textbox(left, top, width, height)
        title_frame = title_box.text_frame
        title_frame.word_wrap = True
        
        p = title_frame.paragraphs[0]
        p.text = title
        p.font.name = self.fonts['heading']
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = RGBColor(*self.colors['primary'])
        p.alignment = PP_ALIGN.CENTER
        
        # Subtitle
        if subtitle:
            sub_left = Inches(1)
            sub_top = Inches(4.2)
            sub_width = Inches(10)
            sub_height = Inches(1)
            
            sub_box = slide.shapes.add_textbox(sub_left, sub_top, sub_width, sub_height)
            sub_frame = sub_box.text_frame
            
            p = sub_frame.paragraphs[0]
            p.text = subtitle
            p.font.name = self.fonts['subheading']
            p.font.size = Pt(18)
            p.font.color.rgb = RGBColor(*self.colors['light_text'])
            p.alignment = PP_ALIGN.CENTER
        
        # Date and issue info
        date_text = datetime.now().strftime("%B %d, %Y")
        issue_text = f"Volume 1 • Issue {datetime.now().month}"
        
        date_left = Inches(1)
        date_top = Inches(5.5)
        date_width = Inches(10)
        date_height = Inches(0.5)
        
        date_box = slide.shapes.add_textbox(date_left, date_top, date_width, date_height)
        date_frame = date_box.text_frame
        
        p = date_frame.paragraphs[0]
        p.text = f"{date_text} • {issue_text}"
        p.font.name = self.fonts['body']
        p.font.size = Pt(14)
        p.font.color.rgb = RGBColor(*self.colors['accent'])
        p.alignment = PP_ALIGN.CENTER
        
        # Add logo if exists
        if self.logo_path and self.logo_path.exists():
            try:
                left = prs.slide_width - Inches(2)
                top = Inches(0.3)
                slide.shapes.add_picture(
                    str(self.logo_path),
                    left, top, 
                    height=Inches(1)
                )
            except Exception as e:
                logger.warning(f"Could not add logo: {str(e)}")
    
    def _add_section_header(self, prs: Presentation, title: str) -> None:
        """Add a newsletter section header slide"""
        slide_layout = prs.slide_layouts[self.slide_layouts['blank']]
        slide = prs.slides.add_slide(slide_layout)
        
        # Background
        self._create_gradient_background(slide, self.colors['primary'], self.colors['secondary'])
        
        # Section title
        left = Inches(1)
        top = Inches(3)
        width = Inches(10)
        height = Inches(2)
        
        title_box = slide.shapes.add_textbox(left, top, width, height)
        title_frame = title_box.text_frame
        
        p = title_frame.paragraphs[0]
        p.text = title.upper()
        p.font.name = self.fonts['heading']
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.color.rgb = RGBColor(*self.colors['white'])
        p.alignment = PP_ALIGN.CENTER
        
        # Decorative line
        line_left = Inches(4)
        line_top = Inches(5.2)
        line_width = Inches(4)
        line_height = Inches(0.1)
        
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, line_left, line_top, line_width, line_height
        )
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(*self.colors['accent'])
        line.line.fill.background()
    
    def _add_newsletter_content_slide(self, prs: Presentation, insights: List[Dict[str, Any]], 
                                    slide_title: str = "") -> None:
        """Add a newsletter-style content slide with multiple articles"""
        slide_layout = prs.slide_layouts[self.slide_layouts['blank']]
        slide = prs.slides.add_slide(slide_layout)
        
        # Background
        self._create_gradient_background(slide, self.colors['background'], self.colors['white'])
        
        # Header with reduced bottom margin
        issue_info = datetime.now().strftime("%B %Y") + " Newsletter"
        self._add_newsletter_header(slide, slide_title or "Latest Updates", issue_info)
        
        # Calculate available height (total slide height - header - bottom margin)
        available_height = 6.0  # Total available height in inches (from 1.5 to 7.5)
        
        # Add article cards in a grid layout with better spacing
        if len(insights) == 1:
            # Single large article with more vertical space
            self._add_article_card(slide, insights[0], (1.5, 1.5, 10.0, 5.0))
        elif len(insights) == 2:
            # Two side-by-side articles with adjusted spacing
            self._add_article_card(slide, insights[0], (0.5, 1.5, 5.5, 5.5))
            self._add_article_card(slide, insights[1], (6.5, 1.5, 5.5, 5.5))
        else:
            # Multiple articles in grid with adjusted spacing
            # Calculate dynamic positions based on number of articles
            if len(insights) <= 4:
                # 2x2 grid for 3-4 articles
                positions = [
                    (0.8, 1.5, 5.5, 2.5),   # Top left
                    (7.0, 1.5, 5.5, 2.5),   # Top right
                    (0.8, 4.2, 5.5, 2.5),   # Bottom left
                    (7.0, 4.2, 5.5, 2.5),   # Bottom right
                ]
            else:
                # 3x2 grid for 5-6 articles
                positions = [
                    (0.5, 1.5, 4.0, 2.5),   # Row 1, Col 1
                    (4.8, 1.5, 4.0, 2.5),   # Row 1, Col 2
                    (9.1, 1.5, 4.0, 2.5),   # Row 1, Col 3
                    (0.5, 4.2, 4.0, 2.5),   # Row 2, Col 1
                    (4.8, 4.2, 4.0, 2.5),   # Row 2, Col 2
                    (9.1, 4.2, 4.0, 2.5),   # Row 2, Col 3
                ]
            
            for i, insight in enumerate(insights[:6]):  # Max 6 articles per slide
                if i < len(positions):
                    # Truncate content if needed before adding to card
                    if 'content' in insight and len(insight['content']) > 120:
                        insight = insight.copy()  # Don't modify original
                        insight['content'] = insight['content'][:120] + '...'
                    self._add_article_card(slide, insight, positions[i])
    
    def _add_summary_slide(self, prs: Presentation, insights: List[Dict[str, Any]]) -> None:
        """Add a newsletter summary slide with statistics"""
        slide_layout = prs.slide_layouts[self.slide_layouts['blank']]
        slide = prs.slides.add_slide(slide_layout)
        
        # Background
        self._create_gradient_background(slide, self.colors['background'], self.colors['white'])
        
        # Header
        self._add_newsletter_header(slide, "Newsletter Summary", "Key Statistics & Insights")
        
        # Add statistics visualization
        self._add_stats_visualization(slide, insights)
        
        # Key highlights box
        highlights_left = Inches(1)
        highlights_top = Inches(5.5)
        highlights_width = Inches(10)
        highlights_height = Inches(1.5)
        
        highlights_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            highlights_left, highlights_top, highlights_width, highlights_height
        )
        
        highlights_box.fill.solid()
        highlights_box.fill.fore_color.rgb = RGBColor(*self.colors['accent'])
        highlights_box.line.fill.background()
        
        # Highlights text
        highlights_frame = slide.shapes.add_textbox(
            Inches(1.3), Inches(5.8), Inches(9.4), Inches(0.9)
        ).text_frame
        
        p = highlights_frame.paragraphs[0]
        total_articles = len(insights)
        categories = len(set(cat for insight in insights for cat in insight.get('categories', [])))
        p.text = f"📊 This newsletter covers {total_articles} articles across {categories} categories"
        p.font.name = self.fonts['subheading']
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = RGBColor(*self.colors['white'])
        p.alignment = PP_ALIGN.CENTER
    
    def generate_presentation(self, insights: List[Dict[str, Any]], 
                            output_path: Optional[str] = None) -> str:
        """
        Generate a newsletter-style PowerPoint presentation from insights
        
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