import os
import sys
from pathlib import Path
import streamlit.web.cli as stcli

def main():
    """Main entry point for the application."""
    # Add the project root to the Python path
    project_root = Path(__file__).parent
    sys.path.append(str(project_root))
    
    # Set environment variables
    os.environ["STREAMLIT_SERVER_PORT"] = "8501"
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    
    # Create necessary directories
    (project_root / "data" / "raw_articles").mkdir(parents=True, exist_ok=True)
    (project_root / "data" / "processed").mkdir(parents=True, exist_ok=True)
    (project_root / "data" / "logos").mkdir(parents=True, exist_ok=True)
    (project_root / "output").mkdir(parents=True, exist_ok=True)
    
    # Check if spaCy model is installed
    try:
        import spacy
        nlp = spacy.load("en_core_web_sm")  # Try loading the small model first
    except OSError:
        print("Downloading spaCy model...")
        import subprocess
        try:
            subprocess.run(["python", "-m", "spacy", "download", "en_core_web_sm"], check=True)
            subprocess.run(["python", "-m", "spacy", "download", "en_core_web_lg"], check=True)
        except subprocess.CalledProcessError as e:
            print(f"Error downloading spaCy models: {e}")
            print("Falling back to small model...")
            try:
                subprocess.run(["python", "-m", "spacy", "download", "en_core_web_sm"], check=True)
            except subprocess.CalledProcessError:
                print("Could not download spaCy models. Some NLP features may not work.")
    
    # Run the Streamlit app
    streamlit_app_path = str(project_root / "streamlit_app" / "app.py")
    sys.argv = ["streamlit", "run", streamlit_app_path, "--server.port=8501", "--server.address=0.0.0.0"]
    sys.exit(stcli.main())

if __name__ == "__main__":
    main()
