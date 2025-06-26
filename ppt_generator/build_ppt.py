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
        
        # Modern newsletter color palette
        self.colors = {
            'primary': (41, 84, 144),      # Deep blue
            'secondary': (72, 133, 237),   # Bright blue
            'accent': (255, 87, 51),       # Coral red
            'success': (46, 204, 113),     # Green
            'warning': (241, 196, 15),     # Yellow
            'background': (248, 249, 250), # Light gray
            'white': (255, 255, 255),      # White
            'dark': (33, 37, 41),          # Dark gray
            'light_text': (108, 117, 125), # Light gray text
            'border': (206, 212, 218)      # Border gray
        }
        
        # Typography settings
        self.fonts = {
            'heading': 'Montserrat',
            'subheading': 'Open Sans',
            'body': 'Segoe UI',
            'accent': 'Georgia'
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
    
    def _create_gradient_background(self, slide, color1: tuple, color2: tuple) -> None:
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
        fill.solid()
        fill.fore_color.rgb = RGBColor(*color1)
        
        # Send to back
        bg_shape.z_order = 0
    
    def _add_decorative_header_bar(self, slide, color: tuple = None) -> None:
        """Add a decorative header bar"""
        if not color:
            color = self.colors['accent']
        
        # Add header bar
        left = 0
        top = 0
        width = slide.parent.slide_width
        height = Inches(0.2)
        
        header_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, left, top, width, height
        )
        
        fill = header_bar.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*color)
        
        # Remove border
        header_bar.line.fill.background()
    
    def _add_newsletter_header(self, slide, title: str, issue_info: str = "") -> None:
        """Add a professional newsletter header"""
        # Add decorative top bar
        self._add_decorative_header_bar(slide)
        
        # Newsletter title
        left = Inches(0.5)
        top = Inches(0.3)
        width = Inches(8)
        height = Inches(1)
        
        title_box = slide.shapes.add_textbox(left, top, width, height)
        title_frame = title_box.text_frame
        title_frame.word_wrap = True
        
        p = title_frame.paragraphs[0]
        p.text = title
        p.font.name = self.fonts['heading']
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = RGBColor(*self.colors['primary'])
        p.alignment = PP_ALIGN.LEFT
        
        # Issue info (date, volume, etc.)
        if issue_info:
            left = Inches(8.5)
            top = Inches(0.4)
            width = Inches(3)
            height = Inches(0.8)
            
            info_box = slide.shapes.add_textbox(left, top, width, height)
            info_frame = info_box.text_frame
            
            p = info_frame.paragraphs[0]
            p.text = issue_info
            p.font.name = self.fonts['body']
            p.font.size = Pt(12)
            p.font.color.rgb = RGBColor(*self.colors['light_text'])
            p.alignment = PP_ALIGN.RIGHT
    
    def _add_article_card(self, slide, article: Dict[str, Any], 
                         position: Tuple[float, float, float, float]) -> None:
        """Add a card-style article layout"""
        left, top, width, height = position
        
        # Card background
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, 
            Inches(left), Inches(top), 
            Inches(width), Inches(height)
        )
        
        # Card styling
        fill = card.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*self.colors['white'])
        
        # Card border
        line = card.line
        line.color.rgb = RGBColor(*self.colors['border'])
        line.width = Pt(1)
        
        # Add shadow effect (simulated with another shape)
        shadow = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(left + 0.05), Inches(top + 0.05),
            Inches(width), Inches(height)
        )
        shadow.fill.solid()
        shadow.fill.fore_color.rgb = RGBColor(220, 220, 220)
        shadow.line.fill.background()
        shadow.z_order = card.z_order - 1
        
        # Category tag
        categories = article.get('categories', ['News'])
        if categories:
            tag_left = Inches(left + 0.2)
            tag_top = Inches(top + 0.2)
            tag_width = Inches(1.5)
            tag_height = Inches(0.3)
            
            tag = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                tag_left, tag_top, tag_width, tag_height
            )
            
            tag.fill.solid()
            tag.fill.fore_color.rgb = RGBColor(*self.colors['secondary'])
            tag.line.fill.background()
            
            # Tag text
            tag_frame = tag.text_frame
            tag_frame.margin_left = Inches(0.1)
            tag_frame.margin_right = Inches(0.1)
            tag_frame.margin_top = Inches(0.05)
            tag_frame.margin_bottom = Inches(0.05)
            
            p = tag_frame.paragraphs[0]
            p.text = categories[0].upper()
            p.font.name = self.fonts['body']
            p.font.size = Pt(8)
            p.font.bold = True
            p.font.color.rgb = RGBColor(*self.colors['white'])
            p.alignment = PP_ALIGN.CENTER
        
        # Article title
        title_left = Inches(left + 0.2)
        title_top = Inches(top + 0.6)
        title_width = Inches(width - 0.4)
        title_height = Inches(0.8)
        
        title_box = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
        title_frame = title_box.text_frame
        title_frame.word_wrap = True
        
        p = title_frame.paragraphs[0]
        p.text = article.get('title', 'Article Title')
        p.font.name = self.fonts['subheading']
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = RGBColor(*self.colors['dark'])
        
        # Article content preview
        content_left = Inches(left + 0.2)
        content_top = Inches(top + 1.5)
        content_width = Inches(width - 0.4)
        content_height = Inches(height - 2.2)
        
        content_box = slide.shapes.add_textbox(content_left, content_top, content_width, content_height)
        content_frame = content_box.text_frame
        content_frame.word_wrap = True
        
        content = article.get('content', '')
        preview = content[:150] + "..." if len(content) > 150 else content
        
        p = content_frame.paragraphs[0]
        p.text = preview
        p.font.name = self.fonts['body']
        p.font.size = Pt(11)
        p.font.color.rgb = RGBColor(*self.colors['dark'])
        p.space_after = Pt(8)
        
        # Source and date
        source_left = Inches(left + 0.2)
        source_top = Inches(top + height - 0.5)
        source_width = Inches(width - 0.4)
        source_height = Inches(0.3)
        
        source_box = slide.shapes.add_textbox(source_left, source_top, source_width, source_height)
        source_frame = source_box.text_frame
        
        p = source_frame.paragraphs[0]
        source = article.get('source', 'Unknown Source')
        date = article.get('date', '')
        if date:
            try:
                date = datetime.strptime(date, '%Y-%m-%d').strftime('%b %d, %Y')
            except:
                pass
        
        p.text = f"{source} • {date}" if date else source
        p.font.name = self.fonts['body']
        p.font.size = Pt(9)
        p.font.color.rgb = RGBColor(*self.colors['light_text'])
        p.font.italic = True
    
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
        
        # Header
        issue_info = datetime.now().strftime("%B %Y") + " Newsletter"
        self._add_newsletter_header(slide, slide_title or "Latest Updates", issue_info)
        
        # Add article cards in a grid layout
        if len(insights) == 1:
            # Single large article
            self._add_article_card(slide, insights[0], (1.5, 1.5, 9, 4.5))
        elif len(insights) == 2:
            # Two side-by-side articles
            self._add_article_card(slide, insights[0], (0.5, 1.5, 5.5, 4.5))
            self._add_article_card(slide, insights[1], (6.5, 1.5, 5.5, 4.5))
        else:
            # Multiple articles in grid
            positions = [
                (0.5, 1.5, 3.8, 2.8),   # Top left
                (4.6, 1.5, 3.8, 2.8),   # Top center
                (8.7, 1.5, 3.8, 2.8),   # Top right
                (0.5, 4.5, 3.8, 2.8),   # Bottom left
                (4.6, 4.5, 3.8, 2.8),   # Bottom center
                (8.7, 4.5, 3.8, 2.8),   # Bottom right
            ]
            
            for i, insight in enumerate(insights[:6]):  # Max 6 articles per slide
                if i < len(positions):
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