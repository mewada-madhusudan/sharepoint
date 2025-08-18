from flask import Blueprint, request, jsonify, current_app, send_file
from app.sharepoint_client import SharePointClient
import pandas as pd
import io
import json
from datetime import datetime

api_bp = Blueprint('api', __name__)

@api_bp.route('/data', methods=['GET'])
def get_data():
    """Get paginated data with filtering and sorting"""
    try:
        page = int(request.args.get('page', 1))
        page_size = int(request.args.get('pageSize', 100))
        filters = request.args.get('filters')
        sort_field = request.args.get('sortField')
        sort_order = request.args.get('sortOrder', 'asc')
        
        if filters:
            filters = json.loads(filters)
        
        client = SharePointClient()
        data = client.get_list_items(
            page=page,
            page_size=page_size,
            filters=filters,
            sort_field=sort_field,
            sort_order=sort_order
        )
        
        return jsonify(data)
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@api_bp.route('/fields', methods=['GET'])
def get_fields():
    """Get list field definitions"""
    try:
        client = SharePointClient()
        fields = client.get_list_fields()
        return jsonify({'fields': fields})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@api_bp.route('/item', methods=['POST'])
def create_item():
    """Create a new item"""
    try:
        data = request.json
        client = SharePointClient()
        item_id = client.create_item(data)
        
        if item_id:
            return jsonify({'success': True, 'id': item_id})
        else:
            return jsonify({'success': False, 'error': 'Failed to create item'}), 400
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@api_bp.route('/item/<int:item_id>', methods=['PUT'])
def update_item(item_id):
    """Update an existing item"""
    try:
        data = request.json
        client = SharePointClient()
        success = client.update_item(item_id, data)
        
        return jsonify({'success': success})
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@api_bp.route('/item/<int:item_id>', methods=['DELETE'])
def delete_item(item_id):
    """Delete an item"""
    try:
        client = SharePointClient()
        success = client.delete_item(item_id)
        
        return jsonify({'success': success})
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@api_bp.route('/bulk', methods=['POST'])
def bulk_operations():
    """Perform bulk operations"""
    try:
        operations = request.json.get('operations', [])
        client = SharePointClient()
        results = client.bulk_update(operations)
        
        return jsonify(results)
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@api_bp.route('/export/<format>', methods=['GET'])
def export_data(format):
    """Export data to Excel or CSV"""
    try:
        client = SharePointClient()
        df = client.export_to_dataframe()
        
        if df.empty:
            return jsonify({'error': 'No data to export'}), 400
        
        # Create file in memory
        output = io.BytesIO()
        
        if format.lower() == 'excel':
            df.to_excel(output, index=False, engine='openpyxl')
            mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            filename = f'sharepoint_data_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        elif format.lower() == 'csv':
            df.to_csv(output, index=False)
            mimetype = 'text/csv'
            filename = f'sharepoint_data_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv'
        else:
            return jsonify({'error': 'Unsupported format'}), 400
        
        output.seek(0)
        
        return send_file(
            output,
            mimetype=mimetype,
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@api_bp.route('/search', methods=['POST'])
def search_data():
    """Global search across all fields"""
    try:
        search_term = request.json.get('searchTerm', '')
        client = SharePointClient()
        
        # Get all data and perform client-side search
        # In production, implement server-side search
        all_data = client.get_list_items(page_size=1000)
        
        if not search_term:
            return jsonify(all_data)
        
        filtered_items = []
        for item in all_data['items']:
            for key, value in item.items():
                if search_term.lower() in str(value).lower():
                    filtered_items.append(item)
                    break
        
        return jsonify({
            'items': filtered_items,
            'total': len(filtered_items),
            'fields': all_data['fields']
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@api_bp.route('/validate', methods=['POST'])
def validate_data():
    """Validate data before saving"""
    try:
        data = request.json
        client = SharePointClient()
        fields = client.get_list_fields()
        
        errors = []
        
        # Validate required fields
        for field in fields:
            if field['required'] and field['name'] not in data:
                errors.append(f"{field['title']} is required")
            
            # Validate field types and formats
            if field['name'] in data:
                value = data[field['name']]
                
                if field['type'] == 'DateTime' and value:
                    try:
                        datetime.fromisoformat(value.replace('Z', '+00:00'))
                    except:
                        errors.append(f"{field['title']} has invalid date format")
                
                elif field['type'] == 'Number' and value is not None:
                    try:
                        float(value)
                    except:
                        errors.append(f"{field['title']} must be a number")
                
                elif field['type'] == 'Choice' and value:
                    if value not in field.get('choices', []):
                        errors.append(f"{field['title']} has invalid choice: {value}")
        
        return jsonify({
            'valid': len(errors) == 0,
            'errors': errors
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500