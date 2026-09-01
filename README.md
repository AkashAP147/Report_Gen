# Report Generator Pro

A streamlined web application that automatically generates examination seating arrangements and dispatch reports directly from an Excel session file.

## Features

- **Intelligent Branch Discovery:** Instantly parses uploaded Excel files to detect available student branches (subjects).
- **Advanced Seating Configuration:** Allows administrators to define precisely which branches should be prioritized on alternating desks to prevent cheating.
- **Direct Multi-File Generation:** Generates 4 strictly formatted `.docx` reports simultaneously.
- **Premium UI:** Features a modern, glassmorphism-inspired dark mode interface with asynchronous drag-and-drop workflows.

## Technology Stack

- **Backend:** Python, Flask, Pandas, Python-Docx
- **Frontend:** Vanilla HTML, CSS (Inter Font), JavaScript
- **Server:** Gunicorn (Production-ready)

## Local Development

### Prerequisites
- Python 3.9+

### Setup

1. Clone the repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the development server:
   ```bash
   python app.py
   ```
4. Access the application at `http://127.0.0.1:5000`

## Deployment

This application is configured to run on modern PaaS platforms like **Render**, **Heroku**, or **Railway**. 

The included `Procfile` uses `gunicorn` to serve the application in a production environment. 

Temporary generation files are safely ignored via `.gitignore` and handled dynamically in memory or local temporary directories.
