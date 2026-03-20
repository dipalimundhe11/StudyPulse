
// Excel Handler for User Data - Using SheetJS Library
const ExcelHandler = {
    // Storage key for Excel file data in localStorage
    STORAGE_KEY: 'userDataExcel',
    
    // Initialize or get existing workbook
    getWorkbook: function() {
        const storedData = localStorage.getItem(this.STORAGE_KEY);
        if (storedData) {
            // Parse base64 string to workbook
            return XLSX.read(storedData, { type: 'base64' });
        } else {
            // Create new workbook with users sheet
            const wb = XLSX.utils.book_new();
            
            // Create users data array with headers
            const usersData = [
                ['Full Name', 'Email', 'Password', 'Action', 'Date/Time']
            ];
            
            // Convert to worksheet and add to workbook
            const ws = XLSX.utils.aoa_to_sheet(usersData);
            XLSX.utils.book_append_sheet(wb, ws, 'Users');
            
            // Create messages sheet
            const messagesData = [
                ['Name', 'Email', 'Phone', 'Subject', 'Message', 'Date/Time']
            ];
            const wsMessages = XLSX.utils.aoa_to_sheet(messagesData);
            XLSX.utils.book_append_sheet(wb, wsMessages, 'Messages');
            
            return wb;
        }
    },
    
    // Save workbook to localStorage
    saveWorkbook: function(wb) {
        // Convert workbook to base64 string
        const excelData = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
        localStorage.setItem(this.STORAGE_KEY, excelData);
    },
    
    // Add new user on signup
    addUser: function(fullName, email, password) {
        try {
            const wb = this.getWorkbook();
            const ws = wb.Sheets['Users'];
            
            // Get existing data
            const existingData = XLSX.utils.sheet_to_json(ws, { header: 1 });
            
            // Add new user row
            const newRow = [fullName, email, password, 'Signup', new Date().toLocaleString()];
            existingData.push(newRow);
            
            // Update worksheet
            const newWs = XLSX.utils.aoa_to_sheet(existingData);
            wb.Sheets['Users'] = newWs;
            
            // Save workbook
            this.saveWorkbook(wb);
            
            console.log('User data saved to Excel successfully');
            return true;
        } catch (error) {
            console.error('Error saving user to Excel:', error);
            return false;
        }
    },
    
    // Record login activity
    recordLogin: function(fullName, email) {
        try {
            const wb = this.getWorkbook();
            const ws = wb.Sheets['Users'];
            
            // Get existing data
            const existingData = XLSX.utils.sheet_to_json(ws, { header: 1 });
            
            // Find user and add login record
            let userFound = false;
            for (let i = 1; i < existingData.length; i++) {
                if (existingData[i][1] === email) {
                    // Add new login record
                    const loginRow = [fullName, email, '', 'Login', new Date().toLocaleString()];
                    existingData.push(loginRow);
                    userFound = true;
                    break;
                }
            }
            
            if (userFound) {
                // Update worksheet
                const newWs = XLSX.utils.aoa_to_sheet(existingData);
                wb.Sheets['Users'] = newWs;
                
                // Save workbook
                this.saveWorkbook(wb);
                
                console.log('Login recorded in Excel successfully');
            }
            
            return userFound;
        } catch (error) {
            console.error('Error recording login in Excel:', error);
            return false;
        }
    },
    
    // Save contact message
    saveMessage: function(name, email, phone, subject, message) {
        try {
            const wb = this.getWorkbook();
            
            // Check if Messages sheet exists, if not create it
            let ws;
            if (wb.Sheets['Messages']) {
                ws = wb.Sheets['Messages'];
                // Get existing data
                const existingData = XLSX.utils.sheet_to_json(ws, { header: 1 });
                
                // Add new message row
                const newRow = [name, email, phone, subject, message, new Date().toLocaleString()];
                existingData.push(newRow);
                
                // Update worksheet
                ws = XLSX.utils.aoa_to_sheet(existingData);
                wb.Sheets['Messages'] = ws;
            } else {
                // Create new messages sheet
                const messagesData = [
                    ['Name', 'Email', 'Phone', 'Subject', 'Message', 'Date/Time'],
                    [name, email, phone, subject, message, new Date().toLocaleString()]
                ];
                ws = XLSX.utils.aoa_to_sheet(messagesData);
                XLSX.utils.book_append_sheet(wb, ws, 'Messages');
            }
            
            // Save workbook
            this.saveWorkbook(wb);
            
            console.log('Message saved to Excel successfully');
            return true;
        } catch (error) {
            console.error('Error saving message to Excel:', error);
            return false;
        }
    },
    
    // Export Excel file (download)
    downloadExcel: function(filename = 'UserData.xlsx') {
        try {
            const wb = this.getWorkbook();
            XLSX.writeFile(wb, filename);
            console.log('Excel file downloaded successfully:', filename);
            return true;
        } catch (error) {
            console.error('Error downloading Excel file:', error);
            return false;
        }
    }
};

// Make it globally available
window.ExcelHandler = ExcelHandler;

