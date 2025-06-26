# 📰 Hybrid Sector Intelligence Newsletter Generator

A Python application that generates professional-quality sector-specific newsletters and exports them as PowerPoint presentations. The app uses open-source APIs for data collection and local models for NLP tasks with a Streamlit-based GUI.

## 🚀 Features

- **Sector-Specific Analysis**: Generate newsletters for various industry sectors
- **Hybrid Data Collection**: Combines multiple data sources including news APIs and web scraping
- **Local NLP Processing**: Uses local models for privacy and efficiency
- **Professional Output**: Generates PowerPoint presentations with clean, structured content
- **Customizable**: Configure data sources, models, and output templates

## 🛠 Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/sector-intelligence-newsletter.git
   cd sector-intelligence-newsletter
   ```

2. Create and activate a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Set up environment variables:
   Create a `.env` file in the project root with your API keys:
   ```
   NEWS_API_KEY=your_news_api_key
   GNEWS_API_KEY=your_gnews_api_key
   ```

## 🚀 Usage

1. Start the application:
   ```bash
   python main.py
   ```

2. Open your web browser and navigate to:
   ```
   http://localhost:8501
   ```

3. Select your desired sector, date range, and insight types

4. Click "Generate Newsletter" and wait for the processing to complete

5. Download the generated PowerPoint presentation

## 🏗 Project Structure

```
newsletter_generator/
├── data/                   # Data storage
│   ├── raw_articles/       # Raw articles from APIs/scrapers
│   ├── processed/          # Processed and cleaned data
│   └── logos/              # Company logos for presentations
├── models/                 # ML models and embeddings
├── scrapers/               # Data collection modules
├── nlp_pipeline/           # Text processing and analysis
├── ppt_generator/          # PowerPoint generation
├── streamlit_app/          # Web interface
│   └── app.py             # Main Streamlit application
├── .env.example           # Example environment variables
├── config.yaml            # Application configuration
├── main.py                # Application entry point
└── README.md              # This file
```

## 🤖 Technologies Used

- **Frontend**: Streamlit
- **NLP**: spaCy, Transformers, Sentence-Transformers
- **Data Processing**: Pandas, NumPy
- **Visualization**: Matplotlib, Plotly
- **Presentation**: python-pptx
- **APIs**: NewsAPI, GNews, SEC EDGAR

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built with ❤️ using open-source software
- Special thanks to the open-source community for their contributions
