// SharePoint Excel Interface - Main Application
class SharePointExcelApp {
    constructor() {
        this.grid = null;
        this.gridOptions = null;
        this.fields = [];
        this.pendingChanges = new Map();
        this.isConnected = false;
        this.currentData = [];
        
        this.initializeGrid();
        this.setupEventListeners();
        this.loadInitialData();
    }

    initializeGrid() {
        const gridDiv = document.querySelector('#myGrid');
        
        this.gridOptions = {
            columnDefs: [],
            rowData: [],
            defaultColDef: {
                editable: true,
                resizable: true,
                sortable: true,
                filter: true,
                floatingFilter: true,
                cellEditor: 'agTextCellEditor'
            },
            enableRangeSelection: true,
            enableFillHandle: true,
            enableCellChangeFlash: true,
            rowSelection: 'multiple',
            suppressRowClickSelection: false,
            animateRows: true,
            pagination: true,
            paginationPageSize: 100,
            sideBar: {
                toolPanels: ['filters', 'columns']
            },
            onCellEditingStopped: (event) => this.onCellChanged(event),
            onSelectionChanged: () => this.updateSelectionCount(),
            onGridReady: (params) => {
                this.gridApi = params.api;
                this.gridColumnApi = params.columnApi;
                params.api.sizeColumnsToFit();
            },
            // Excel-like keyboard navigation
            navigateToNextCell: (params) => {
                const suggestedNextCell = params.nextCellPosition;
                const KEY_UP = 38;
                const KEY_DOWN = 40;
                const KEY_LEFT = 37;
                const KEY_RIGHT = 39;

                switch (params.key) {
                    case KEY_DOWN:
                        return {rowIndex: suggestedNextCell.rowIndex + 1, column: suggestedNextCell.column};
                    case KEY_UP:
                        return {rowIndex: suggestedNextCell.rowIndex - 1, column: suggestedNextCell.column};
                    case KEY_LEFT:
                        return {rowIndex: suggestedNextCell.rowIndex, column: this.gridColumnApi.getDisplayedColAfter(suggestedNextCell.column)};
                    case KEY_RIGHT:
                        return {rowIndex: suggestedNextCell.rowIndex, column: this.gridColumnApi.getDisplayedColBefore(suggestedNextCell.column)};
                    default:
                        return suggestedNextCell;
                }
            }
        };

        this.grid = new agGrid.Grid(gridDiv, this.gridOptions);
    }

    setupEventListeners() {
        // Toolbar buttons
        document.getElementById('addRowBtn').addEventListener('click', () => this.addNewRow());
        document.getElementById('deleteRowBtn').addEventListener('click', () => this.deleteSelectedRows());
        document.getElementById('saveChangesBtn').addEventListener('click', () => this.saveAllChanges());
        document.getElementById('refreshBtn').addEventListener('click', () => this.refreshData());
        
        // Export buttons
        document.getElementById('exportExcel').addEventListener('click', () => this.exportData('excel'));
        document.getElementById('exportCsv').addEventListener('click', () => this.exportData('csv'));
        
        // Search
        document.getElementById('globalSearch').addEventListener('input', (e) => this.performGlobalSearch(e.target.value));
        document.getElementById('clearSearch').addEventListener('click', () => this.clearSearch());
        
        // Keyboard shortcuts
        document.addEventListener('keydown', (e) => this.handleKeyboardShortcuts(e));
        
        // Modal events
        document.getElementById('saveRowBtn').addEventListener('click', () => this.saveRowFromModal());
        document.getElementById('confirmActionBtn').addEventListener('click', () => this.executeConfirmedAction());
    }

    async loadInitialData() {
        this.showLoading(true);
        
        try {
            // First get field definitions
            const fieldsResponse = await fetch('/api/fields');
            const fieldsData = await fieldsResponse.json();
            
            if (fieldsData.error) {
                throw new Error(fieldsData.error);
            }
            
            this.fields = fieldsData.fields || [];
            this.setupColumns();
            
            // Then get data
            const dataResponse = await fetch('/api/data');
            const data = await dataResponse.json();
            
            if (data.error) {
                throw new Error(data.error);
            }
            
            this.currentData = data.items || [];
            this.gridApi.setRowData(this.currentData);
            
            this.updateStatusInfo(data.total || 0);
            this.isConnected = true;
            this.showError(null);
            
        } catch (error) {
            console.error('Error loading data:', error);
            this.showError(error.message);
            this.isConnected = false;
        } finally {
            this.showLoading(false);
        }
    }

    setupColumns() {
        const columns = [
            {
                headerName: '',
                checkboxSelection: true,
                headerCheckboxSelection: true,
                width: 50,
                pinned: 'left',
                editable: false,
                filter: false,
                sortable: false
            }
        ];

        this.fields.forEach(field => {
            const colDef = {
                field: field.name,
                headerName: field.title,
                editable: true,
                filter: true,
                sortable: true
            };

            // Configure based on field type
            switch (field.type) {
                case 'DateTime':
                    colDef.cellEditor = 'agDateStringCellEditor';
                    colDef.valueFormatter = (params) => {
                        if (params.value) {
                            return new Date(params.value).toLocaleString();
                        }
                        return '';
                    };
                    break;
                case 'Number':
                    colDef.cellEditor = 'agNumberCellEditor';
                    colDef.filter = 'agNumberColumnFilter';
                    break;
                case 'Choice':
                    colDef.cellEditor = 'agSelectCellEditor';
                    colDef.cellEditorParams = {
                        values: field.choices || []
                    };
                    break;
                case 'Boolean':
                    colDef.cellRenderer = 'agCheckboxCellRenderer';
                    colDef.cellEditor = 'agCheckboxCellEditor';
                    break;
                default:
                    colDef.cellEditor = 'agTextCellEditor';
            }

            // Mark required fields
            if (field.required) {
                colDef.headerName += ' *';
                colDef.cellStyle = {
                    'border-left': '3px solid #dc3545'
                };
            }

            columns.push(colDef);
        });

        this.gridApi.setColumnDefs(columns);
    }

    onCellChanged(event) {
        const rowId = event.data.ID || 'new_' + Date.now();
        const fieldName = event.colDef.field;
        const newValue = event.newValue;
        const oldValue = event.oldValue;

        if (newValue !== oldValue) {
            // Track pending changes
            if (!this.pendingChanges.has(rowId)) {
                this.pendingChanges.set(rowId, { original: {...event.data}, changes: {} });
            }
            
            this.pendingChanges.get(rowId).changes[fieldName] = newValue;
            
            // Visual feedback
            event.node.setRowClass('ag-row-pending');
            
            this.updatePendingChangesCount();
            this.enableSaveButton(true);
        }
    }

    async addNewRow() {
        const modal = new bootstrap.Modal(document.getElementById('editModal'));
        document.getElementById('editModalTitle').textContent = 'Add New Row';
        
        this.buildFormFields();
        modal.show();
    }

    buildFormFields(rowData = {}) {
        const container = document.getElementById('formFields');
        container.innerHTML = '';

        this.fields.forEach(field => {
            const div = document.createElement('div');
            div.className = 'mb-3';

            const label = document.createElement('label');
            label.className = 'form-label';
            label.textContent = field.title + (field.required ? ' *' : '');
            
            let input;
            
            switch (field.type) {
                case 'Choice':
                    input = document.createElement('select');
                    input.className = 'form-select';
                    
                    const defaultOption = document.createElement('option');
                    defaultOption.value = '';
                    defaultOption.textContent = 'Select...';
                    input.appendChild(defaultOption);
                    
                    field.choices?.forEach(choice => {
                        const option = document.createElement('option');
                        option.value = choice;
                        option.textContent = choice;
                        input.appendChild(option);
                    });
                    break;
                    
                case 'Boolean':
                    input = document.createElement('input');
                    input.type = 'checkbox';
                    input.className = 'form-check-input';
                    break;
                    
                case 'Number':
                    input = document.createElement('input');
                    input.type = 'number';
                    input.className = 'form-control';
                    break;
                    
                case 'DateTime':
                    input = document.createElement('input');
                    input.type = 'datetime-local';
                    input.className = 'form-control';
                    break;
                    
                default:
                    input = document.createElement('input');
                    input.type = 'text';
                    input.className = 'form-control';
            }
            
            input.id = field.name;
            input.name = field.name;
            
            if (rowData[field.name] !== undefined) {
                if (field.type === 'Boolean') {
                    input.checked = rowData[field.name];
                } else {
                    input.value = rowData[field.name];
                }
            }
            
            div.appendChild(label);
            div.appendChild(input);
            container.appendChild(div);
        });
    }

    async saveRowFromModal() {
        const formData = new FormData(document.getElementById('editForm'));
        const data = {};
        
        // Convert form data to object
        for (let [key, value] of formData.entries()) {
            const field = this.fields.find(f => f.name === key);
            
            if (field) {
                if (field.type === 'Boolean') {
                    data[key] = document.getElementById(key).checked;
                } else if (field.type === 'Number') {
                    data[key] = value ? parseFloat(value) : null;
                } else {
                    data[key] = value || null;
                }
            }
        }
        
        // Validate required fields
        const errors = this.validateRowData(data);
        if (errors.length > 0) {
            this.showToast('Please fix the following errors:\n' + errors.join('\n'), 'error');
            return;
        }
        
        try {
            const response = await fetch('/api/item', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(data)
            });
            
            const result = await response.json();
            
            if (result.success) {
                this.showToast('Row added successfully', 'success');
                bootstrap.Modal.getInstance(document.getElementById('editModal')).hide();
                this.refreshData();
            } else {
                throw new Error(result.error || 'Failed to add row');
            }
        } catch (error) {
            this.showToast('Error adding row: ' + error.message, 'error');
        }
    }

    validateRowData(data) {
        const errors = [];
        
        this.fields.forEach(field => {
            if (field.required && (!data[field.name] || data[field.name] === '')) {
                errors.push(`${field.title} is required`);
            }
        });
        
        return errors;
    }

    async deleteSelectedRows() {
        const selectedRows = this.gridApi.getSelectedRows();
        
        if (selectedRows.length === 0) {
            this.showToast('Please select rows to delete', 'error');
            return;
        }
        
        const modal = new bootstrap.Modal(document.getElementById('confirmModal'));
        document.getElementById('confirmMessage').textContent = 
            `Are you sure you want to delete ${selectedRows.length} row(s)? This action cannot be undone.`;
        
        this.pendingConfirmAction = async () => {
            try {
                for (const row of selectedRows) {
                    const response = await fetch(`/api/item/${row.ID}`, {
                        method: 'DELETE'
                    });
                    
                    if (!response.ok) {
                        throw new Error(`Failed to delete row ${row.ID}`);
                    }
                }
                
                this.showToast(`${selectedRows.length} row(s) deleted successfully`, 'success');
                this.refreshData();
            } catch (error) {
                this.showToast('Error deleting rows: ' + error.message, 'error');
            }
        };
        
        modal.show();
    }

    async saveAllChanges() {
        if (this.pendingChanges.size === 0) {
            this.showToast('No changes to save', 'info');
            return;
        }
        
        const operations = [];
        
        for (const [rowId, change] of this.pendingChanges) {
            if (rowId.startsWith('new_')) {
                operations.push({
                    action: 'create',
                    data: {...change.original, ...change.changes}
                });
            } else {
                operations.push({
                    action: 'update',
                    id: rowId,
                    data: change.changes
                });
            }
        }
        
        try {
            const response = await fetch('/api/bulk', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ operations })
            });
            
            const result = await response.json();
            
            if (result.success) {
                this.showToast('All changes saved successfully', 'success');
                this.pendingChanges.clear();
                this.updatePendingChangesCount();
                this.enableSaveButton(false);
                this.refreshData();
            } else {
                throw new Error('Some operations failed: ' + JSON.stringify(result.errors));
            }
        } catch (error) {
            this.showToast('Error saving changes: ' + error.message, 'error');
        }
    }

    async refreshData() {
        await this.loadInitialData();
        this.pendingChanges.clear();
        this.updatePendingChangesCount();
        this.enableSaveButton(false);
    }

    async exportData(format) {
        try {
            const response = await fetch(`/api/export/${format}`);
            
            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                
                const contentDisposition = response.headers.get('content-disposition');
                let filename = `sharepoint_data_${new Date().toISOString().slice(0,10)}.${format}`;
                
                if (contentDisposition) {
                    const match = contentDisposition.match(/filename="(.+)"/);
                    if (match) filename = match[1];
                }
                
                a.href = url;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
                
                this.showToast('Export completed successfully', 'success');
            } else {
                throw new Error('Export failed');
            }
        } catch (error) {
            this.showToast('Error exporting data: ' + error.message, 'error');
        }
    }

    performGlobalSearch(searchTerm) {
        if (searchTerm) {
            this.gridApi.setQuickFilter(searchTerm);
        } else {
            this.gridApi.setQuickFilter('');
        }
    }

    clearSearch() {
        document.getElementById('globalSearch').value = '';
        this.gridApi.setQuickFilter('');
    }

    handleKeyboardShortcuts(event) {
        if (event.ctrlKey || event.metaKey) {
            switch (event.key) {
                case 's':
                    event.preventDefault();
                    this.saveAllChanges();
                    break;
                case 'n':
                    event.preventDefault();
                    this.addNewRow();
                    break;
                case 'r':
                    event.preventDefault();
                    this.refreshData();
                    break;
                case 'f':
                    event.preventDefault();
                    document.getElementById('globalSearch').focus();
                    break;
            }
        }
        
        if (event.key === 'Delete') {
            const selectedRows = this.gridApi.getSelectedRows();
            if (selectedRows.length > 0) {
                this.deleteSelectedRows();
            }
        }
    }

    updateSelectionCount() {
        const selectedRows = this.gridApi.getSelectedRows();
        document.getElementById('selectedRows').textContent = `Selected: ${selectedRows.length}`;
        
        const deleteBtn = document.getElementById('deleteRowBtn');
        deleteBtn.disabled = selectedRows.length === 0;
    }

    updateStatusInfo(total) {
        document.getElementById('totalRows').textContent = `Total: ${total}`;
        document.getElementById('lastUpdated').textContent = 
            `Last updated: ${new Date().toLocaleTimeString()}`;
    }

    updatePendingChangesCount() {
        document.getElementById('pendingChanges').textContent = 
            `Pending: ${this.pendingChanges.size}`;
    }

    enableSaveButton(enabled) {
        document.getElementById('saveChangesBtn').disabled = !enabled;
    }

    executeConfirmedAction() {
        if (this.pendingConfirmAction) {
            this.pendingConfirmAction();
            this.pendingConfirmAction = null;
        }
        bootstrap.Modal.getInstance(document.getElementById('confirmModal')).hide();
    }

    showLoading(show) {
        const spinner = document.getElementById('loadingSpinner');
        const grid = document.getElementById('myGrid');
        
        if (show) {
            spinner.style.display = 'block';
            grid.style.display = 'none';
        } else {
            spinner.style.display = 'none';
            grid.style.display = 'block';
        }
    }

    showError(message) {
        const errorAlert = document.getElementById('errorAlert');
        const errorMessage = document.getElementById('errorMessage');
        
        if (message) {
            errorMessage.textContent = message;
            errorAlert.style.display = 'block';
        } else {
            errorAlert.style.display = 'none';
        }
    }

    showToast(message, type = 'info') {
        const toastElement = type === 'error' ? 
            document.getElementById('errorToast') : 
            document.getElementById('successToast');
        
        const messageElement = type === 'error' ? 
            document.getElementById('errorToastMessage') : 
            document.getElementById('successMessage');
        
        messageElement.textContent = message;
        
        const toast = new bootstrap.Toast(toastElement);
        toast.show();
    }
}

// Initialize the application when the DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    new SharePointExcelApp();
});