import os
from dotenv import load_dotenv

basedir = os.path.abspath(os.path.dirname(__file__))
load_dotenv(os.path.join(basedir, '.env'))

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'sharepoint-excel-app-secret-key'
    SQLALCHEMY_DATABASE_URI = os.environ.get('DATABASE_URL') or \
        'sqlite:///' + os.path.join(basedir, 'app.db')
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    
    # SharePoint Configuration
    SHAREPOINT_URL = os.environ.get('SHAREPOINT_URL') or 'https://your-tenant.sharepoint.com/sites/your-site'
    SHAREPOINT_USERNAME = os.environ.get('SHAREPOINT_USERNAME') or ''
    SHAREPOINT_PASSWORD = os.environ.get('SHAREPOINT_PASSWORD') or ''
    SHAREPOINT_LIST_NAME = os.environ.get('SHAREPOINT_LIST_NAME') or 'Your List Name'
    
    # App Configuration
    ROWS_PER_PAGE = int(os.environ.get('ROWS_PER_PAGE') or 100)
    MAX_EXPORT_ROWS = int(os.environ.get('MAX_EXPORT_ROWS') or 10000)