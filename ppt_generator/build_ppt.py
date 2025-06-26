from typing import List, Dict, Any, Optional, Tuple
from pathlib import Path
from datetime import datetime
import logging
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
import os
import json

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class PPTGenerator:
    """Generate PowerPoint presentations from article insights"""
    
    def __init__(self, config: dict):
        """Initialize the PPT generator with configuration"""
        self.config = config
        self.template_path = Path(config.get('template_path', ''))
        self.output_dir = Path(config.get('output_dir', 'output'))
        self.logo_path = Path(config.get('logo_path', 'assets/logo.png'))
        
        # Create output directory if it doesn't exist
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Default colors
        self.colors = {
            'background': (255, 255, 255),  # White
            'text': (0, 0, 0),             # Black
            'accent': (0, 84, 147),         # Blue
            'highlight': (255, 204, 0),     # Yellow
            'light_gray': (240, 240, 240)   # Light gray
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
    
    def _apply_slide_theme(self, slide, theme: str = 'default') -> None:
        """Apply a theme to a slide"""
        # This is a placeholder - in a real implementation, you would apply
        # specific formatting based on the theme
        pass
    
    def _add_title_slide(self, prs: Presentation, title: str, subtitle: str = "") -> None:
        """Add a title slide to the presentation"""
        slide_layout = prs.slide_layouts[self.slide_layouts['title']]
        slide = prs.slides.add_slide(slide_layout)
        
        # Add title
        title_shape = slide.shapes.title
        title_shape.text = title
        
        # Add subtitle if provided
        if subtitle:
            subtitle_shape = slide.placeholders[1]
            subtitle_shape.text = subtitle
        
        # Add date
        date_shape = slide.shapes.add_textbox(
            Inches(0.5), 
            prs.slide_height - Inches(1),
            Inches(4),
            Inches(0.5)
        )
        date_frame = date_shape.text_frame
        date_frame.text = datetime.now().strftime("%B %d, %Y")
        
        # Add logo if exists
        if self.logo_path and self.logo_path.exists():
            try:
                left = prs.slide_width - Inches(1.5)
                top = Inches(0.5)
                slide.shapes.add_picture(
                    str(self.logo_path),
                    left, top, 
                    height=Inches(0.8)
                )
            except Exception as e:
                logger.warning(f"Could not add logo: {str(e)}")
    
    def _add_section_header(self, prs: Presentation, title: str) -> None:
        """Add a section header slide"""
        slide_layout = prs.slide_layouts[self.slide_layouts['section_header']]
        slide = prs.slides.add_slide(slide_layout)
        
        # Set title
        title_shape = slide.shapes.title
        title_shape.text = title
        
        # Apply theme
        self._apply_slide_theme(slide, 'section')
    
    def _add_insight_slide(self, prs: Presentation, insight: Dict[str, Any]) -> None:
        """Add a slide for a single insight"""
        slide_layout = prs.slide_layouts[self.slide_layouts['content']]
        slide = prs.slides.add_slide(slide_layout)
        
        # Set title
        title_shape = slide.shapes.title
        title_shape.text = insight.get('title', 'Insight')
        
        # Add content
        content_shape = slide.placeholders[1]
        text_frame = content_shape.text_frame
        
        # Add source and date if available
        source = insight.get('source', 'Unknown')
        date = insight.get('date', '')
        if date:
            try:
                date = datetime.strptime(date, '%Y-%m-%d').strftime('%B %d, %Y')
            except:
                pass
        
        # Add source and date
        p = text_frame.add_paragraph()
        p.text = f"Source: {source}"
        if date:
            p.text += f" | {date}"
        p.font.size = Pt(12)
        p.font.italic = True
        p.font.color.rgb = RGBColor(100, 100, 100)
        
        # Add content
        content = insight.get('content', '')
        if content:
            p = text_frame.add_paragraph()
            p.text = content
            p.space_after = Pt(12)
        
        # Add key points if available
        key_points = insight.get('key_points', [])
        if key_points:
            p = text_frame.add_paragraph()
            p.text = "Key Points:"
            p.font.bold = True
            
            for point in key_points:
                p = text_frame.add_paragraph()
                p.text = f"• {point}"
                p.level = 1
        
        # Add footer with categories if available
        categories = insight.get('categories', [])
        if categories:
            footer = slide.shapes.add_textbox(
                Inches(0.5),
                prs.slide_height - Inches(0.7),
                prs.slide_width - Inches(1),
                Inches(0.5)
            )
            footer_frame = footer.text_frame
            p = footer_frame.paragraphs[0]
            p.text = "Categories: " + ", ".join(categories)
            p.font.size = Pt(10)
            p.font.color.rgb = RGBColor(100, 100, 100)
    
    def _add_comparison_slide(self, prs: Presentation, insights: List[Dict[str, Any]]) -> None:
        """Add a slide comparing multiple insights"""
        if len(insights) < 2:
            return self._add_insight_slide(prs, insights[0] if insights else {})
        
        slide_layout = prs.slide_layouts[self.slide_layouts['comparison']]
        slide = prs.slides.add_slide(slide_layout)
        
        # Set title
        title_shape = slide.shapes.title
        title_shape.text = "Comparison"
        
        # Add content to left and right placeholders
        left_shape = slide.placeholders[1]
        right_shape = slide.placeholders[2]
        
        # Add first insight to left
        left_frame = left_shape.text_frame
        left_frame.text = insights[0].get('title', 'Insight 1')
        p = left_frame.add_paragraph()
        p.text = insights[0].get('content', '')[:200] + "..."
        
        # Add second insight to right
        right_frame = right_shape.text_frame
        right_frame.text = insights[1].get('title', 'Insight 2')
        p = right_frame.add_paragraph()
        p.text = insights[1].get('content', '')[:200] + "..."
    
    def _add_summary_slide(self, prs: Presentation, insights: List[Dict[str, Any]]) -> None:
        """Add a summary slide with key takeaways"""
        slide_layout = prs.slide_layouts[self.slide_layouts['title_only']]
        slide = prs.slides.add_slide(slide_layout)
        
        # Add title
        title_shape = slide.shapes.title
        title_shape.text = "Key Takeaways"
        
        # Add content
        left = Inches(1)
        top = Inches(1.5)
        width = prs.slide_width - Inches(2)
        height = prs.slide_height - Inches(2.5)
        
        content = slide.shapes.add_textbox(left, top, width, height)
        text_frame = content.text_frame
        
        # Group insights by category
        categories = {}
        for insight in insights:
            for category in insight.get('categories', ['Other']):
                if category not in categories:
                    categories[category] = []
                categories[category].append(insight)
        
        # Add content by category
        for category, items in categories.items():
            p = text_frame.add_paragraph()
            p.text = category.upper()
            p.font.bold = True
            p.space_after = Pt(6)
            
            for item in items[:3]:  # Limit to top 3 per category
                p = text_frame.add_paragraph()
                p.text = f"• {item.get('title', '')}"
                p.level = 1
                p.space_after = Pt(2)
            
            text_frame.add_paragraph()  # Add space between categories
    
    def generate_presentation(self, insights: List[Dict[str, Any]], 
                            output_path: Optional[str] = None) -> str:
        """
        Generate a PowerPoint presentation from insights
        
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
        
        # Set presentation properties
        prs.slide_width = 12193280  # 16:9 aspect ratio
        prs.slide_height = 6858000
        
        # Add title slide
        sector = insights[0].get('sector', 'Sector')
        self._add_title_slide(
            prs, 
            title=f"{sector} Intelligence Report",
            subtitle="Monthly Insights and Analysis"
        )
        
        # Group insights by category
        categorized = {}
        for insight in insights:
            cats = insight.get('categories', ['Other'])
            for cat in cats:
                if cat not in categorized:
                    categorized[cat] = []
                categorized[cat].append(insight)
        
        # Add slides for each category
        for category, items in categorized.items():
            self._add_section_header(prs, category)
            
            # Add slides for each insight in this category
            for i in range(0, len(items), 2):
                batch = items[i:i+2]
                if len(batch) == 2:
                    self._add_comparison_slide(prs, batch)
                else:
                    self._add_insight_slide(prs, batch[0])
        
        # Add summary slide
        self._add_summary_slide(prs, insights)
        
        # Save the presentation
        if not output_path:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = self.output_dir / f"{sector.lower().replace(' ', '_')}_report_{timestamp}.pptx"
        else:
            output_path = Path(output_path)
        
        prs.save(str(output_path))
        logger.info(f"Presentation saved to {output_path}")
        
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
    """Generate sample insights for testing"""
    return [
        {
            "title": "Tech Giant Acquires AI Startup",
            "content": "A major technology company has announced the acquisition of a leading AI startup for an undisclosed amount. The deal is expected to close in Q3 2023.",
            "source": "Tech News Daily",
            "date": "2023-06-15",
            "categories": ["Mergers & Acquisitions", "Artificial Intelligence"],
            "sector": "Technology",
            "key_points": [
                "Acquisition expected to enhance AI capabilities",
                "Deal terms not disclosed",
                "Closing expected in Q3 2023"
            ]
        },
        {
            "title": "New Renewable Energy Initiative Launched",
            "content": "A coalition of energy companies has launched a new initiative to increase renewable energy production by 50% over the next decade.",
            "source": "Energy Today",
            "date": "2023-06-10",
            "categories": ["Renewable Energy", "Sustainability"],
            "sector": "Energy",
            "key_points": [
                "Aim to increase renewable energy production by 50%",
                "Multi-company collaboration",
                "10-year timeline for implementation"
            ]
        }
    ]
