#!/usr/bin/env python3
"""
SharePoint Excel Interface - Main application entry point
"""
import os
from app import create_app

app = create_app()

if __name__ == '__main__':
    # Get port from environment variable or default to 5000
    port = int(os.environ.get('PORT', 5000))
    
    # Run the application
    app.run(
        debug=True,
        host='0.0.0.0',
        port=port
    )