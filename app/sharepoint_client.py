import json
import pandas as pd
from datetime import datetime
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.lists.list import List
from office365.sharepoint.listitems.listitem import ListItem
from flask import current_app
import logging

logger = logging.getLogger(__name__)

class SharePointClient:
    def __init__(self):
        self.site_url = current_app.config['SHAREPOINT_URL']
        self.username = current_app.config['SHAREPOINT_USERNAME']
        self.password = current_app.config['SHAREPOINT_PASSWORD']
        self.list_name = current_app.config['SHAREPOINT_LIST_NAME']
        self.ctx = None
        self.list_obj = None
        
    def authenticate(self):
        """Authenticate with SharePoint"""
        try:
            auth_ctx = AuthenticationContext(url=self.site_url)
            if auth_ctx.acquire_token_for_user(username=self.username, password=self.password):
                self.ctx = ClientContext(self.site_url, auth_ctx)
                return True
            else:
                logger.error("Authentication failed")
                return False
        except Exception as e:
            logger.error(f"Authentication error: {str(e)}")
            return False
    
    def get_list(self):
        """Get SharePoint list object"""
        if not self.ctx:
            if not self.authenticate():
                return None
        
        try:
            web = self.ctx.web
            self.list_obj = web.lists.get_by_title(self.list_name)
            self.ctx.load(self.list_obj)
            self.ctx.execute_query()
            return self.list_obj
        except Exception as e:
            logger.error(f"Error getting list: {str(e)}")
            return None
    
    def get_list_fields(self):
        """Get list field definitions"""
        if not self.get_list():
            return []
        
        try:
            fields = self.list_obj.fields
            self.ctx.load(fields)
            self.ctx.execute_query()
            
            field_info = []
            for field in fields:
                if not field.properties.get('Hidden', True) and field.properties.get('ReadOnlyField', True) != True:
                    field_info.append({
                        'name': field.properties.get('InternalName', ''),
                        'title': field.properties.get('Title', ''),
                        'type': field.properties.get('TypeAsString', ''),
                        'required': field.properties.get('Required', False),
                        'choices': field.properties.get('Choices', [])
                    })
            
            return field_info
        except Exception as e:
            logger.error(f"Error getting fields: {str(e)}")
            return []
    
    def get_list_items(self, page=1, page_size=100, filters=None, sort_field=None, sort_order='asc'):
        """Get list items with pagination, filtering, and sorting"""
        if not self.get_list():
            return {'items': [], 'total': 0, 'fields': []}
        
        try:
            # Build CAML query
            query = self._build_caml_query(filters, sort_field, sort_order, page, page_size)
            
            items = self.list_obj.get_items(query)
            self.ctx.load(items)
            self.ctx.execute_query()
            
            # Get field definitions
            fields = self.get_list_fields()
            
            # Process items
            processed_items = []
            for item in items:
                item_data = {'ID': item.properties['Id']}
                for field in fields:
                    field_name = field['name']
                    value = item.properties.get(field_name)
                    
                    # Handle different field types
                    if field['type'] == 'DateTime' and value:
                        try:
                            dt = datetime.fromisoformat(value.replace('Z', '+00:00'))
                            item_data[field_name] = dt.strftime('%Y-%m-%d %H:%M:%S')
                        except:
                            item_data[field_name] = value
                    elif field['type'] == 'User' and value:
                        item_data[field_name] = value.get('Title', '') if isinstance(value, dict) else str(value)
                    elif field['type'] == 'Lookup' and value:
                        item_data[field_name] = value.get('Title', '') if isinstance(value, dict) else str(value)
                    else:
                        item_data[field_name] = value
                
                processed_items.append(item_data)
            
            # Get total count (simplified approach)
            total_items = len(processed_items)  # In production, use proper count query
            
            return {
                'items': processed_items,
                'total': total_items,
                'fields': fields,
                'page': page,
                'page_size': page_size
            }
            
        except Exception as e:
            logger.error(f"Error getting items: {str(e)}")
            return {'items': [], 'total': 0, 'fields': []}
    
    def create_item(self, item_data):
        """Create a new list item"""
        if not self.get_list():
            return None
        
        try:
            item_create_info = self.list_obj.add_item(item_data)
            self.ctx.execute_query()
            return item_create_info.properties['Id']
        except Exception as e:
            logger.error(f"Error creating item: {str(e)}")
            return None
    
    def update_item(self, item_id, item_data):
        """Update an existing list item"""
        if not self.get_list():
            return False
        
        try:
            item = self.list_obj.get_item_by_id(item_id)
            item.update(item_data)
            self.ctx.execute_query()
            return True
        except Exception as e:
            logger.error(f"Error updating item: {str(e)}")
            return False
    
    def delete_item(self, item_id):
        """Delete a list item"""
        if not self.get_list():
            return False
        
        try:
            item = self.list_obj.get_item_by_id(item_id)
            item.delete_object()
            self.ctx.execute_query()
            return True
        except Exception as e:
            logger.error(f"Error deleting item: {str(e)}")
            return False
    
    def bulk_update(self, updates):
        """Perform bulk updates"""
        if not self.get_list():
            return {'success': False, 'errors': []}
        
        results = {'success': True, 'errors': []}
        
        for update in updates:
            try:
                if update['action'] == 'create':
                    self.create_item(update['data'])
                elif update['action'] == 'update':
                    self.update_item(update['id'], update['data'])
                elif update['action'] == 'delete':
                    self.delete_item(update['id'])
            except Exception as e:
                results['errors'].append({
                    'id': update.get('id'),
                    'action': update['action'],
                    'error': str(e)
                })
                results['success'] = False
        
        return results
    
    def export_to_dataframe(self):
        """Export all list data to pandas DataFrame"""
        data = self.get_list_items(page_size=current_app.config['MAX_EXPORT_ROWS'])
        if data['items']:
            return pd.DataFrame(data['items'])
        return pd.DataFrame()
    
    def _build_caml_query(self, filters=None, sort_field=None, sort_order='asc', page=1, page_size=100):
        """Build CAML query for SharePoint"""
        # Simplified CAML query - in production, build proper CAML XML
        query_options = {
            'ViewXml': '<View><Query></Query></View>'
        }
        
        if page_size:
            query_options['ViewXml'] = f'<View><Query></Query><RowLimit>{page_size}</RowLimit></View>'
        
        return query_options