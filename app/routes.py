from flask import Blueprint, render_template, current_app

main_bp = Blueprint('main', __name__)

@main_bp.route('/')
def index():
    """Main application page"""
    return render_template('index.html', 
                         app_name="SharePoint Excel Interface",
                         list_name=current_app.config['SHAREPOINT_LIST_NAME'])

@main_bp.route('/health')
def health():
    """Health check endpoint"""
    return {'status': 'healthy', 'app': 'SharePoint Excel Interface'}