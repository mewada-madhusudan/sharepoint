import json
import pandas as pd
from datetime import datetime
from flask import current_app
import logging

# Office365 imports for SharePoint Online
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

# Shareplum imports for on-premises SharePoint
from shareplum import Site, Office365
from shareplum.site import Version
import requests
from requests_ntlm import HttpNtlmAuth

logger = logging.getLogger(__name__)

class SharePointClient:
    def __init__(self):
        self.site_url = current_app.config['SHAREPOINT_URL']
        self.username = current_app.config['SHAREPOINT_USERNAME']
        self.password = current_app.config['SHAREPOINT_PASSWORD']
        self.list_name = current_app.config['SHAREPOINT_LIST_NAME']
        self.ctx = None
        self.list_obj = None
        self.site = None
        self.sp_list = None
        self.is_onprem = self._is_onpremise_url(self.site_url)
        
    def _is_onpremise_url(self, url):
        """Check if the URL is on-premise SharePoint (not sharepoint.com)"""
        return 'sharepoint.com' not in url.lower()
        
    def authenticate(self):
        """Authenticate with SharePoint (On-premise using Shareplum or Online using Office365)"""
        try:
            if self.is_onprem:
                return self._authenticate_onprem()
            else:
                return self._authenticate_online()
                
        except Exception as e:
            logger.error(f"Authentication error: {str(e)}")
            return False
    
    def _authenticate_onprem(self):
        """Authenticate with on-premise SharePoint using Shareplum"""
        try:
            # Try NTLM authentication first
            auth = HttpNtlmAuth(self.username, self.password)
            
            # Create site connection with Shareplum
            self.site = Site(
                site_url=self.site_url,
                auth=auth,
                version=Version.v2007  # Specify SharePoint 2007
            )
            
            # Test connection by getting site information
            site_info = self.site.site_info
            logger.info(f"Successfully connected to on-premise SharePoint 2007: {site_info.get('Title', 'Unknown')}")
            
            # Get the list
            self.sp_list = self.site.List(self.list_name)
            
            return True
            
        except Exception as e:
            logger.error(f"On-premise authentication failed: {str(e)}")
            
            # Try with basic authentication as fallback
            try:
                import requests
                from requests.auth import HTTPBasicAuth
                
                basic_auth = HTTPBasicAuth(self.username, self.password)
                self.site = Site(
                    site_url=self.site_url,
                    auth=basic_auth,
                    version=Version.v2007
                )
                
                site_info = self.site.site_info
                logger.info(f"Successfully connected with basic auth to: {site_info.get('Title', 'Unknown')}")
                
                self.sp_list = self.site.List(self.list_name)
                return True
                
            except Exception as basic_error:
                logger.error(f"Basic authentication also failed: {str(basic_error)}")
                return False
    
    def _authenticate_online(self):
        """Authenticate with SharePoint Online using Office365"""
        try:
            from office365.runtime.auth.authentication_context import AuthenticationContext
            auth_ctx = AuthenticationContext(url=self.site_url)
            if auth_ctx.acquire_token_for_user(username=self.username, password=self.password):
                self.ctx = ClientContext(self.site_url, auth_ctx)
                logger.info("Successfully authenticated with SharePoint Online")
                return True
            else:
                logger.error("SharePoint Online authentication failed")
                return False
        except Exception as e:
            logger.error(f"SharePoint Online authentication error: {str(e)}")
            return False
    
    def get_list(self):
        """Get SharePoint list object"""
        if not self.site and not self.ctx:
            if not self.authenticate():
                return None
        
        if self.is_onprem:
            return self.sp_list
        else:
            try:
                web = self.ctx.web
                self.list_obj = web.lists.get_by_title(self.list_name)
                self.ctx.load(self.list_obj)
                self.ctx.execute_query()
                return self.list_obj
            except Exception as e:
                logger.error(f"Error getting online list: {str(e)}")
                return None
    
    def get_list_fields(self):
        """Get list field definitions"""
        if self.is_onprem:
            return self._get_onprem_fields()
        else:
            return self._get_online_fields()
    
    def _get_onprem_fields(self):
        """Get field definitions for on-premise SharePoint using Shareplum"""
        if not self.sp_list:
            if not self.authenticate():
                return []
        
        try:
            # For Shareplum, we need to get a sample item to understand the fields
            # or use the site's field definitions
            sample_items = self.sp_list.get_list_items(fields=['Title'], rows=1)
            
            if sample_items:
                # Get field names from the first item
                sample_item = sample_items[0]
                field_info = []
                
                # Standard SharePoint fields
                standard_fields = {
                    'ID': {'title': 'ID', 'type': 'Counter', 'required': False},
                    'Title': {'title': 'Title', 'type': 'Text', 'required': True},
                    'Created': {'title': 'Created', 'type': 'DateTime', 'required': False},
                    'Modified': {'title': 'Modified', 'type': 'DateTime', 'required': False},
                    'Author': {'title': 'Created By', 'type': 'User', 'required': False},
                    'Editor': {'title': 'Modified By', 'type': 'User', 'required': False}
                }
                
                # Add standard fields first
                for field_name, field_props in standard_fields.items():
                    field_info.append({
                        'name': field_name,
                        'title': field_props['title'],
                        'type': field_props['type'],
                        'required': field_props['required'],
                        'choices': []
                    })
                
                # Add custom fields found in the sample item
                for field_name in sample_item.keys():
                    if field_name not in standard_fields:
                        # Try to guess field type based on value
                        field_type = self._guess_field_type(sample_item[field_name])
                        field_info.append({
                            'name': field_name,
                            'title': field_name.replace('_', ' ').title(),
                            'type': field_type,
                            'required': False,
                            'choices': []
                        })
                
                return field_info
            else:
                # No items found, return basic fields
                return self._get_default_fields()
            
        except Exception as e:
            logger.error(f"Error getting on-premise fields: {str(e)}")
            return self._get_default_fields()
    
    def _get_default_fields(self):
        """Return default SharePoint fields"""
        return [
            {'name': 'ID', 'title': 'ID', 'type': 'Counter', 'required': False, 'choices': []},
            {'name': 'Title', 'title': 'Title', 'type': 'Text', 'required': True, 'choices': []},
            {'name': 'Created', 'title': 'Created', 'type': 'DateTime', 'required': False, 'choices': []},
            {'name': 'Modified', 'title': 'Modified', 'type': 'DateTime', 'required': False, 'choices': []},
            {'name': 'Author', 'title': 'Created By', 'type': 'User', 'required': False, 'choices': []},
            {'name': 'Editor', 'title': 'Modified By', 'type': 'User', 'required': False, 'choices': []}
        ]
    
    def _guess_field_type(self, value):
        """Guess field type based on value"""
        if value is None:
            return 'Text'
        
        value_str = str(value)
        
        # Check for datetime patterns
        datetime_patterns = [
            r'\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}',
            r'\d{1,2}/\d{1,2}/\d{4}',
            r'\d{4}-\d{2}-\d{2}'
        ]
        
        import re
        for pattern in datetime_patterns:
            if re.search(pattern, value_str):
                return 'DateTime'
        
        # Check for numbers
        try:
            float(value_str)
            if '.' in value_str:
                return 'Number'
            else:
                return 'Integer'
        except ValueError:
            pass
        
        # Check for boolean
        if value_str.lower() in ['true', 'false', '1', '0']:
            return 'Boolean'
        
        # Default to text
        return 'Text'
    
    def _get_online_fields(self):
        """Get field definitions for SharePoint Online"""
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
            logger.error(f"Error getting online fields: {str(e)}")
            return []
    
    def get_list_items(self, page=1, page_size=100, filters=None, sort_field=None, sort_order='asc'):
        """Get list items with pagination, filtering, and sorting"""
        if self.is_onprem:
            return self._get_onprem_items(page, page_size, filters, sort_field, sort_order)
        else:
            return self._get_online_items(page, page_size, filters, sort_field, sort_order)
    
    def _get_onprem_items(self, page=1, page_size=100, filters=None, sort_field=None, sort_order='asc'):
        """Get items from on-premise SharePoint using Shareplum"""
        if not self.sp_list:
            if not self.authenticate():
                return {'items': [], 'total': 0, 'fields': []}
        
        try:
            # Calculate row limit for pagination
            # Shareplum supports rows parameter for limiting results
            max_rows = page * page_size
            
            # Get items using Shareplum
            # You can specify fields to retrieve and number of rows
            items = self.sp_list.get_list_items(rows=max_rows)
            
            # Get field definitions
            fields = self.get_list_fields()
            
            # Process items
            processed_items = []
            for item in items:
                item_data = {}
                
                # Process all fields found in the item
                for key, value in item.items():
                    # Find field definition for proper type handling
                    field_def = next((f for f in fields if f['name'] == key), None)
                    field_type = field_def['type'] if field_def else 'Text'
                    
                    # Handle different field types
                    if field_type in ['DateTime', 'Date'] and value:
                        try:
                            if isinstance(value, str):
                                # Try to parse various date formats
                                date_formats = [
                                    '%Y-%m-%dT%H:%M:%SZ',
                                    '%Y-%m-%dT%H:%M:%S.%fZ',
                                    '%Y-%m-%d %H:%M:%S',
                                    '%m/%d/%Y %H:%M:%S',
                                    '%m/%d/%Y',
                                    '%Y-%m-%d'
                                ]
                                
                                for fmt in date_formats:
                                    try:
                                        dt = datetime.strptime(value, fmt)
                                        item_data[key] = dt.strftime('%Y-%m-%d %H:%M:%S')
                                        break
                                    except ValueError:
                                        continue
                                else:
                                    item_data[key] = value
                            else:
                                item_data[key] = str(value)
                        except:
                            item_data[key] = value
                    elif field_type in ['User', 'Lookup'] and value:
                        # Handle user and lookup fields
                        if isinstance(value, dict):
                            item_data[key] = value.get('Title', str(value))
                        elif isinstance(value, str) and ';#' in value:
                            # SharePoint lookup format: "ID;#Value"
                            parts = value.split(';#')
                            item_data[key] = parts[1] if len(parts) > 1 else parts[0]
                        else:
                            item_data[key] = str(value) if value is not None else ''
                    else:
                        item_data[key] = value
                
                processed_items.append(item_data)
            
            # Apply client-side filtering
            if filters:
                processed_items = self._apply_filters(processed_items, filters)
            
            # Apply client-side sorting
            if sort_field and sort_field in (processed_items[0].keys() if processed_items else []):
                reverse = sort_order.lower() == 'desc'
                try:
                    processed_items.sort(key=lambda x: x.get(sort_field, ''), reverse=reverse)
                except TypeError:
                    # Handle mixed types by converting to string
                    processed_items.sort(key=lambda x: str(x.get(sort_field, '')), reverse=reverse)
            
            # Apply pagination
            total_items = len(processed_items)
            start_idx = (page - 1) * page_size
            end_idx = start_idx + page_size
            paginated_items = processed_items[start_idx:end_idx]
            
            return {
                'items': paginated_items,
                'total': total_items,
                'fields': fields,
                'page': page,
                'page_size': page_size
            }
            
        except Exception as e:
            logger.error(f"Error getting on-premise items: {str(e)}")
            return {'items': [], 'total': 0, 'fields': []}
    
    def _get_online_items(self, page=1, page_size=100, filters=None, sort_field=None, sort_order='asc'):
        """Get items from SharePoint Online"""
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
            total_items = len(processed_items)
            
            return {
                'items': processed_items,
                'total': total_items,
                'fields': fields,
                'page': page,
                'page_size': page_size
            }
            
        except Exception as e:
            logger.error(f"Error getting online items: {str(e)}")
            return {'items': [], 'total': 0, 'fields': []}
    
    def _apply_filters(self, items, filters):
        """Apply client-side filters to items"""
        if not filters:
            return items
        
        filtered_items = []
        for item in items:
            match = True
            for field_name, filter_value in filters.items():
                item_value = str(item.get(field_name, '')).lower()
                if str(filter_value).lower() not in item_value:
                    match = False
                    break
            if match:
                filtered_items.append(item)
        
        return filtered_items
    
    def create_item(self, item_data):
        """Create a new list item"""
        if self.is_onprem:
            return self._create_onprem_item(item_data)
        else:
            return self._create_online_item(item_data)
    
    def _create_onprem_item(self, item_data):
        """Create item in on-premise SharePoint using Shareplum"""
        if not self.sp_list:
            if not self.authenticate():
                return None
        
        try:
            result = self.sp_list.update_list_items(data=[item_data], kind='New')
            if result and len(result) > 0:
                return result[0].get('ID')
            return None
        except Exception as e:
            logger.error(f"Error creating on-premise item: {str(e)}")
            return None
    
    def _create_online_item(self, item_data):
        """Create item in SharePoint Online"""
        if not self.get_list():
            return None
        
        try:
            item_create_info = self.list_obj.add_item(item_data)
            self.ctx.execute_query()
            return item_create_info.properties['Id']
        except Exception as e:
            logger.error(f"Error creating online item: {str(e)}")
            return None
    
    def update_item(self, item_id, item_data):
        """Update an existing list item"""
        if self.is_onprem:
            return self._update_onprem_item(item_id, item_data)
        else:
            return self._update_online_item(item_id, item_data)
    
    def _update_onprem_item(self, item_id, item_data):
        """Update item in on-premise SharePoint using Shareplum"""
        if not self.sp_list:
            if not self.authenticate():
                return False
        
        try:
            # Add ID to the data for update
            update_data = dict(item_data)
            update_data['ID'] = item_id
            
            result = self.sp_list.update_list_items(data=[update_data], kind='Update')
            return result is not None
        except Exception as e:
            logger.error(f"Error updating on-premise item: {str(e)}")
            return False
    
    def _update_online_item(self, item_id, item_data):
        """Update item in SharePoint Online"""
        if not self.get_list():
            return False
        
        try:
            item = self.list_obj.get_item_by_id(item_id)
            item.update(item_data)
            self.ctx.execute_query()
            return True
        except Exception as e:
            logger.error(f"Error updating online item: {str(e)}")
            return False
    
    def delete_item(self, item_id):
        """Delete a list item"""
        if self.is_onprem:
            return self._delete_onprem_item(item_id)
        else:
            return self._delete_online_item(item_id)
    
    def _delete_onprem_item(self, item_id):
        """Delete item from on-premise SharePoint using Shareplum"""
        if not self.sp_list:
            if not self.authenticate():
                return False
        
        try:
            result = self.sp_list.update_list_items(data=[{'ID': item_id}], kind='Delete')
            return result is not None
        except Exception as e:
            logger.error(f"Error deleting on-premise item: {str(e)}")
            return False
    
    def _delete_online_item(self, item_id):
        """Delete item from SharePoint Online"""
        if not self.get_list():
            return False
        
        try:
            item = self.list_obj.get_item_by_id(item_id)
            item.delete_object()
            self.ctx.execute_query()
            return True
        except Exception as e:
            logger.error(f"Error deleting online item: {str(e)}")
            return False
    
    def bulk_update(self, updates):
        """Perform bulk updates"""
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
        """Build CAML query for SharePoint Online"""
        # Simplified CAML query - in production, build proper CAML XML
        query_options = {
            'ViewXml': '<View><Query></Query></View>'
        }
        
        if page_size:
            query_options['ViewXml'] = f'<View><Query></Query><RowLimit>{page_size}</RowLimit></View>'
        
        return query_options
